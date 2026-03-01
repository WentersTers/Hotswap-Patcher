namespace PAIcomPatcher.Core;

/// <summary>
/// The C# source that gets compiled by Roslyn at patch-time and then
/// grafted into the target assembly as a new type.
///
/// This is kept as a plain string so it can be edited without needing to
/// rebuild the patcher itself – you can also externalise it to a .cs file
/// on disk if you prefer to edit it with IDE tooling.
///
/// The compiled type will be inserted into the target assembly and called
/// from the patched init method (via HotSwap.StartWatcher).
/// </summary>
public static class HotSwapTemplate
{
    /// <summary>
    /// Modify this source to implement your actual hot-swap behaviour.
    /// It is compiled against the host's .NET runtime references at patch
    /// time, so you have access to the full BCL.
    /// </summary>
    public const string SourceCode = """
        using System;
        using System.IO;
        using System.Threading;

        /// <summary>
        /// Hot-swap runtime injected into PAIcom.exe.
        ///
        /// Responsibilities:
        ///   1. Watch command_input.txt for changes.
        ///   2. Read the written command token.
        ///   3. Look it up in commands.txt (format: TOKEN=GAME_COMMAND).
        ///   4. Route the resolved command to the game's own dispatch logic.
        ///   5. Log commands that were dispatched (for debugging).
        /// </summary>
        public static class HotSwapRuntime
        {
            // ── Configuration ─────────────────────────────────────────────
            // Use the directory that contains the running exe so paths are
            // correct regardless of the process working directory at startup.
            private static readonly string BaseDir =
                System.IO.Path.GetDirectoryName(
                    System.Reflection.Assembly.GetExecutingAssembly().Location)
                ?? AppDomain.CurrentDomain.BaseDirectory;

            private static string InputFile    => System.IO.Path.Combine(BaseDir, "command_input.txt");
            private static string CommandsFile => System.IO.Path.Combine(BaseDir, "custom-commands", "commands.txt");
            private static string LogFile      => System.IO.Path.Combine(BaseDir, "hotswap.log");

            // ── State ──────────────────────────────────────────────────────
            private static FileSystemWatcher _watcher;
            private static int _started = 0;
            private static Action<string> _dispatchDelegate;
            private static DateTime _lastRead = DateTime.MinValue;
            private static readonly object _lock = new object();

            // Reflection-located speech engine (found after game finishes init)
            private static object _speechEngine;   // SpeechRecognitionEngine or SpeechRecognizer
            private static System.Reflection.MethodInfo _emulateMethod;
            private static System.Reflection.MethodInfo _emulateAsyncMethod;

            // Direct animation method (bypass speech engine)
            private static System.Reflection.MethodInfo _animationMethod;
            private static bool _animationMethodSearched = false;

            // Game's command-dispatch method (dictionary-based lookup)
            private static System.Reflection.MethodInfo _commandHandlerMethod;
            
            // Stores the main Form so we can invoke on it
            private static System.Windows.Forms.Form _mainForm;

            // PictureBox controls discovered from the form for SHOW/HIDE fallback
            private static System.Windows.Forms.Control[] _pictureBoxes;

            // Methods to stop/restart async recognition for thread-safety
            private static System.Reflection.MethodInfo _recognizeAsyncCancelMethod;
            private static System.Reflection.MethodInfo _recognizeAsyncMethod;

            // ── Entry point ────────────────────────────────────────────────

            /// <summary>
            /// Called once from the patched init method.
            /// Idempotent – calls after the first are ignored.
            /// </summary>
            public static void StartWatcher()
            {
                if (Interlocked.CompareExchange(ref _started, 1, 0) != 0)
                    return;

                try
                {
                    Log("HotSwapRuntime.StartWatcher() initialising …");
                    EnsureInputFile();
                    StartFSW();
                    StartPollingFallback();

                    // Give the game time to finish its own init, then locate
                    // the SpeechRecognitionEngine/SpeechRecognizer via reflection.
                    var t = new System.Threading.Thread(DelayedEngineSearch)
                    {
                        IsBackground = true,
                        Name = "HotSwapEngineSearch"
                    };
                    t.Start();

                    Log("FileSystemWatcher active.  Write a phrase to command_input.txt to dispatch.");
                }
                catch (Exception ex)
                {
                    Log("[ERROR] StartWatcher failed: " + ex);
                }
            }

            /// <summary>
            /// Register the game's command-dispatch delegate so we can call it.
            /// The patcher can optionally wire this up; if not set, we fall
            /// back to a no-op and log the unresolved command.
            /// </summary>
            public static void RegisterDispatchDelegate(Action<string> dispatch)
            {
                _dispatchDelegate = dispatch;
            }

            /// <summary>
            /// Called by the patched command handler on every dispatch so we
            /// can log command activity.
            /// </summary>
            public static void OnCommandDispatched(string command)
            {
                if (command != null)
                    Log("[CMD] " + command);
            }

            // ── FSW watcher ───────────────────────────────────────────────

            private static void StartFSW()
            {
                var dir  = Path.GetFullPath(".");
                _watcher = new FileSystemWatcher(dir, InputFile)
                {
                    NotifyFilter        = NotifyFilters.LastWrite | NotifyFilters.Size,
                    IncludeSubdirectories = false,
                    EnableRaisingEvents = true,
                };
                _watcher.Changed += OnInputFileChanged;
                _watcher.Error   += (_, e) => Log($"[FSW error] {e.GetException()}");
            }

            private static void OnInputFileChanged(object src, FileSystemEventArgs e)
            {
                ProcessInputFile();
            }

            // ── Polling fallback (for write-lock race) ────────────────────

            private static void StartPollingFallback()
            {
                var t = new Thread(PollLoop)
                {
                    IsBackground = true,
                    Name         = "HotSwapPoller",
                };
                t.Start();
            }

            private static void PollLoop()
            {
                while (true)
                {
                    Thread.Sleep(500);
                    try
                    {
                        if (!File.Exists(InputFile)) continue;
                        var wt = File.GetLastWriteTimeUtc(InputFile);
                        if (wt > _lastRead)
                            ProcessInputFile();
                    }
                    catch { /* ignore polling transients */ }
                }
            }

            // ── Core processing ───────────────────────────────────────────

            private static void ProcessInputFile()
            {
                lock (_lock)
                {
                    try
                    {
                        var wt = File.GetLastWriteTimeUtc(InputFile);
                        if (wt <= _lastRead) return; // already handled
                        _lastRead = wt;

                        // Retry loop to handle write-lock
                        string token = "";
                        for (int attempt = 0; attempt < 5; attempt++)
                        {
                            try
                            {
                                token = File.ReadAllText(InputFile).Trim();
                                break;
                            }
                            catch (IOException)
                            {
                                Thread.Sleep(50);
                            }
                        }

                        if (token == null || token.Trim() == string.Empty) return;

                        Log($"[INPUT] token='{token}'");

                        var gameCmd = LookupCommand(token);
                        if (gameCmd == null)
                        {
                            Log("[WARN] No mapping found for token '" + token + "' in " + CommandsFile);
                            return;
                        }

                        DispatchCommand(token, gameCmd);
                    }
                    catch (Exception ex)
                    {
                        Log($"[ERROR] ProcessInputFile: {ex}");
                    }
                }
            }

            // ── Command lookup ────────────────────────────────────────────

            /// <summary>
            /// Reads custom-commands/commands.txt, which has lines of the form:
            ///   hey paicom open youtube (youtube.txt)
            ///
            /// Strips the trailing (*.txt) reference, compares the phrase to the
            /// input token, and returns the referenced filename (without .txt)
            /// as the resolved command name.
            /// Returns null if no match.
            /// </summary>
            private static string LookupCommand(string token)
            {
                if (!File.Exists(CommandsFile)) return null;

                var normalizedToken = token.Trim().ToLowerInvariant();

                foreach (var rawLine in File.ReadAllLines(CommandsFile))
                {
                    var line = rawLine.Trim();
                    if (line.Length == 0 || line.StartsWith("#")) continue;

                    // Extract the (file.txt) reference at the end of the line
                    // Format: "hey paicom do something (file.txt)"
                    var parenOpen  = line.LastIndexOf('(');
                    var parenClose = line.LastIndexOf(')');

                    string phrase;
                    string fileRef;

                    if (parenOpen > 0 && parenClose > parenOpen)
                    {
                        phrase  = line.Substring(0, parenOpen).Trim();
                        fileRef = line.Substring(parenOpen + 1, parenClose - parenOpen - 1).Trim();
                        // Strip .txt extension to get the plain command name
                        if (fileRef.ToLowerInvariant().EndsWith(".txt"))
                            fileRef = fileRef.Substring(0, fileRef.Length - 4);
                    }
                    else
                    {
                        // No parenthesised reference – treat the whole line as phrase
                        phrase  = line;
                        fileRef = line;
                    }

                    if (string.Equals(phrase, normalizedToken, StringComparison.OrdinalIgnoreCase))
                        return fileRef;
                }
                return null;
            }

            // ── Dispatch ──────────────────────────────────────────────────

            /// <summary>
            /// Primary dispatch: try EmulateRecognize so the game's own full
            /// handler fires (animations + audio + URLs).  Falls back to the
            /// local script-runner if the engine has not been found yet.
            /// </summary>
            /// <param name="originalPhrase">Exact phrase as written to command_input.txt</param>
            /// <param name="commandName">Resolved command name (used for script fallback)</param>
            private static void DispatchCommand(string originalPhrase, string commandName)
            {
                Log("[DISPATCH] '" + commandName + "' (phrase: '" + originalPhrase + "')");
                Log("[DISPATCH] Thread: " + System.Threading.Thread.CurrentThread.Name + " (ID " + System.Threading.Thread.CurrentThread.ManagedThreadId + ", IsBackground=" + System.Threading.Thread.CurrentThread.IsBackground + ")");
                Log("[DISPATCH] State: engine=" + (_speechEngine != null) + " emulate=" + (_emulateMethod != null) + " emulateAsync=" + (_emulateAsyncMethod != null) + " animMethod=" + (_animationMethod != null) + " mainForm=" + (_mainForm != null));

                // -- Attempt 1: EmulateRecognizeAsync on the UI thread --
                // SpeechRecognitionEngine is thread-affine; methods must be
                // invoked from the same thread that created the engine (the UI thread).
                if (_emulateAsyncMethod != null && _speechEngine != null && _mainForm != null)
                {
                    try
                    {
                        Log("[EMULATE-ASYNC] Marshalling EmulateRecognizeAsync(\"" + originalPhrase + "\") to UI thread …");
                        var asyncException = new System.Threading.ManualResetEventSlim(false);
                        Exception uiError = null;
                        _mainForm.Invoke(new Action(() =>
                        {
                            try
                            {
                                _emulateAsyncMethod.Invoke(_speechEngine, new object[] { originalPhrase });
                            }
                            catch (Exception ex)
                            {
                                uiError = ex is System.Reflection.TargetInvocationException t ? (t.InnerException ?? ex) : ex;
                            }
                            finally { asyncException.Set(); }
                        }));
                        asyncException.Wait(5000);
                        if (uiError == null)
                        {
                            Log("[EMULATE-ASYNC] OK – game handler will fire asynchronously.");
                            return;
                        }
                        Log("[EMULATE-ASYNC] Failed on UI thread: " + uiError.GetType().Name + ": " + uiError.Message);
                    }
                    catch (Exception ex)
                    {
                        Log("[EMULATE-ASYNC] Invoke failed: " + ex.GetType().Name + ": " + ex.Message);
                    }
                }

                // -- Attempt 2: Cancel async recognition, EmulateRecognize, restart --
                // Nuclear option: temporarily stop async recognition so EmulateRecognize works.
                if (_emulateMethod != null && _speechEngine != null && _mainForm != null
                    && _recognizeAsyncCancelMethod != null && _recognizeAsyncMethod != null)
                {
                    try
                    {
                        Log("[EMULATE-STOP] Stopping async recognition, calling EmulateRecognize, then restarting …");
                        Exception stopError = null;
                        _mainForm.Invoke(new Action(() =>
                        {
                            try
                            {
                                _recognizeAsyncCancelMethod.Invoke(_speechEngine, null);
                                System.Threading.Thread.Sleep(50);
                                _emulateMethod.Invoke(_speechEngine, new object[] { originalPhrase });
                                System.Threading.Thread.Sleep(50);
                                // Restart: RecognizeAsync(RecognizeMode.Multiple) = RecognizeAsync(1)
                                _recognizeAsyncMethod.Invoke(_speechEngine, new object[] { 1 });
                            }
                            catch (Exception ex)
                            {
                                stopError = ex is System.Reflection.TargetInvocationException t ? (t.InnerException ?? ex) : ex;
                            }
                        }));
                        if (stopError == null)
                        {
                            Log("[EMULATE-STOP] OK – recognition restarted.");
                            return;
                        }
                        Log("[EMULATE-STOP] Failed: " + stopError.GetType().Name + ": " + stopError.Message);
                    }
                    catch (Exception ex)
                    {
                        Log("[EMULATE-STOP] Invoke failed: " + ex.GetType().Name + ": " + ex.Message);
                    }
                }

                // -- Attempt 3: Directly invoke the internal animation method on the UI thread --
                if (_animationMethod != null && _mainForm != null)
                {
                    try
                    {
                        var relativePath = "animations/" + commandName + ".txt";
                        Log("[DIRECT-INVOKE] Calling " + _animationMethod.Name + " with \"" + relativePath + "\" on UI thread");
                        
                        _mainForm.BeginInvoke(new Action(() =>
                        {
                            try { _animationMethod.Invoke(_mainForm, new object[] { relativePath }); }
                            catch (Exception uiEx) { Log("[DIRECT-INVOKE] UI thread error: " + uiEx.Message); }
                        }));
                        
                        Log("[DIRECT-INVOKE] Dispatched to UI thread.");
                        return;
                    }
                    catch (Exception ex)
                    {
                        Log("[DIRECT-INVOKE] Failed (" + ex.Message + "). Falling back.");
                    }
                }

                // -- Attempt 4: Try the command handler method (dictionary-based dispatch) --
                if (_commandHandlerMethod != null && _mainForm != null)
                {
                    try
                    {
                        Log("[CMD-HANDLER] Calling " + _commandHandlerMethod.Name + " with \"" + originalPhrase + "\" on UI thread");
                        _mainForm.BeginInvoke(new Action(() =>
                        {
                            try { _commandHandlerMethod.Invoke(_mainForm, new object[] { originalPhrase }); }
                            catch (Exception uiEx) { Log("[CMD-HANDLER] UI thread error: " + uiEx.Message); }
                        }));
                        Log("[CMD-HANDLER] Dispatched to UI thread.");
                        return;
                    }
                    catch (Exception ex)
                    {
                        Log("[CMD-HANDLER] Failed: " + ex.Message);
                    }
                }

                if (_emulateMethod == null && _emulateAsyncMethod == null && _animationMethod == null)
                    Log("[DISPATCH] No engine or animation method available yet; using script fallback.");

                // -- Attempt 5: run the local animation script --
                var scriptPath = System.IO.Path.Combine(BaseDir, "animations", commandName + ".txt");
                if (!File.Exists(scriptPath))
                {
                    Log("[WARN] Animation script not found: " + scriptPath);
                    return;
                }

                Log("[SCRIPT] Running " + scriptPath);
                RunAnimationScript(scriptPath);
            }

            // ── Speech-engine discovery ───────────────────────────────────

            private static void DelayedEngineSearch()
            {
                System.Threading.Thread.Sleep(3000);
                
                int maxAttempts = 30; // 30 × 10 s = 5 min
                for (int attempt = 1; attempt <= maxAttempts; attempt++)
                {
                    if (_emulateMethod != null) return; // already found

                    Log("[ENGINE] Discovery attempt " + attempt + "/" + maxAttempts + " …");
                    try
                    {
                        FindSpeechEngine();
                    }
                    catch (Exception ex)
                    {
                        Log("[ENGINE] Discovery error: " + ex.Message);
                    }

                    if (_emulateMethod != null)
                    {
                        Log("[ENGINE] Engine found after " + attempt + " attempt(s).");
                        return;
                    }

                    if (attempt < maxAttempts)
                        System.Threading.Thread.Sleep(10000); // wait 10 s before retry
                }
                Log("[ENGINE] Speech engine not found after 5 minutes.  Script fallback will be used.");
            }

            private static void FindSpeechEngine()
            {
                // Candidate type names (the obfuscated exe still references the real assembly names)
                var engineTypeNames = new string[]
                {
                    "System.Speech.Recognition.SpeechRecognitionEngine",
                    "System.Speech.Recognition.SpeechRecognizer"
                };

                // Walk every open form and its instance fields (recursively one level)
                System.Windows.Forms.Form[] forms = null;
                try
                {
                    var fc = System.Windows.Forms.Application.OpenForms;
                    forms = new System.Windows.Forms.Form[fc.Count];
                    for (int i = 0; i < fc.Count; i++)
                        forms[i] = fc[i];
                }
                catch (Exception ex)
                {
                    Log("[ENGINE] Cannot enumerate OpenForms: " + ex.Message);
                    return;
                }

                foreach (var form in forms)
                {
                    if (form == null) continue;
                    
                    if (_mainForm == null)
                    {
                        _mainForm = form;
                        FindAnimationMethod(_mainForm);
                    }
                    
                    var found = SearchObjectForEngine(form, engineTypeNames, depth: 0, maxDepth: 3);
                    if (found != null)
                    {
                        StoreEngine(found);
                        return;
                    }
                }

                Log("[ENGINE] Speech engine not found in open forms.");
            }

            private static void FindAnimationMethod(System.Windows.Forms.Form form)
            {
                if (_animationMethod != null || _animationMethodSearched) return;
                _animationMethodSearched = true;
                
                try 
                {
                    var methods = form.GetType().GetMethods(
                        System.Reflection.BindingFlags.Instance | 
                        System.Reflection.BindingFlags.NonPublic | 
                        System.Reflection.BindingFlags.Public);

                    Log("[ENGINE] Scanning " + methods.Length + " methods on form type " + form.GetType().FullName + " …");

                    // Log ALL methods that take a single string param for diagnostics
                    var candidates = new System.Collections.Generic.List<System.Reflection.MethodInfo>();
                    foreach (var m in methods)
                    {
                        var parameters = m.GetParameters();
                        if (parameters.Length == 1 && parameters[0].ParameterType == typeof(string))
                        {
                            var retName = m.ReturnType != null ? m.ReturnType.FullName : "(null)";
                            Log("[ENGINE]   candidate: " + m.Name + "(string) → " + retName);
                            candidates.Add(m);
                        }
                    }
                    Log("[ENGINE] Found " + candidates.Count + " methods with signature ?(string).");

                    // Priority 1: exact Task return
                    foreach (var m in candidates)
                    {
                        var retName = m.ReturnType != null ? m.ReturnType.FullName : "";
                        if (retName == "System.Threading.Tasks.Task")
                        {
                            Log("[ENGINE] ★ Selected animation method (Task return): " + m.Name);
                            _animationMethod = m;
                            break;
                        }
                    }

                    // Priority 2: return type name contains "Task"
                    if (_animationMethod == null)
                    {
                        foreach (var m in candidates)
                        {
                            var retName = m.ReturnType != null ? m.ReturnType.FullName : "";
                            if (retName != null && retName.Contains("Task"))
                            {
                                Log("[ENGINE] ★ Selected animation method (partial Task match): " + m.Name + " → " + retName);
                                _animationMethod = m;
                                break;
                            }
                        }
                    }

                    // Priority 3: async void method (shows as void return but has AsyncStateMachineAttribute)
                    if (_animationMethod == null)
                    {
                        foreach (var m in candidates)
                        {
                            var attrs = m.GetCustomAttributes(false);
                            foreach (var attr in attrs)
                            {
                                if (attr.GetType().FullName != null && attr.GetType().FullName.Contains("AsyncStateMachine"))
                                {
                                    Log("[ENGINE] ★ Selected animation method (async void): " + m.Name);
                                    _animationMethod = m;
                                    break;
                                }
                            }
                            if (_animationMethod != null) break;
                        }
                    }

                    // Also look for the command handler method (void return, takes string,
                    // likely the method that reads from the dictionary and calls the animation method)
                    // We search for methods that reference File.ReadAllLines or dictionary access
                    // For now, collect void(string) methods as candidates for later
                    foreach (var m in candidates)
                    {
                        if (m.ReturnType == typeof(void) && m != _animationMethod)
                        {
                            // Heuristic: look at the method body for dictionary references
                            try
                            {
                                var body = m.GetMethodBody();
                                if (body != null)
                                {
                                    var il = body.GetILAsByteArray();
                                    if (il != null && il.Length > 20 && il.Length < 2000)
                                    {
                                        // Method is non-trivial – could be the command handler
                                        if (_commandHandlerMethod == null)
                                        {
                                            _commandHandlerMethod = m;
                                            Log("[ENGINE] ★ Possible command handler: " + m.Name + " (void, string param, IL size " + il.Length + ")");
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log("[ENGINE] Error finding animation method: " + ex.Message);
                }

                if (_animationMethod == null)
                    Log("[ENGINE] No animation method found after scanning all candidates.");

                // Discover PictureBox controls for SHOW/HIDE script fallback
                DiscoverPictureBoxes(form);
            }

            private static object SearchObjectForEngine(object obj, string[] typeNames, int depth, int maxDepth)
            {
                if (obj == null || depth > maxDepth) return null;

                var type = obj.GetType();

                // Check if this object IS an engine
                foreach (var name in typeNames)
                    if (type.FullName == name)
                        return obj;

                // Scan its instance fields
                var flags = System.Reflection.BindingFlags.Instance |
                            System.Reflection.BindingFlags.NonPublic |
                            System.Reflection.BindingFlags.Public;

                foreach (var field in type.GetFields(flags))
                {
                    try
                    {
                        var value = field.GetValue(obj);
                        if (value == null) continue;

                        var fieldTypeName = value.GetType().FullName;
                        foreach (var name in typeNames)
                        {
                            if (fieldTypeName == name)
                            {
                                Log("[ENGINE] Found " + fieldTypeName + " in field '" + field.Name + "' of " + type.FullName);
                                return value;
                            }
                        }

                        // Recurse into reference types that are likely to hold the engine
                        if (depth < maxDepth && !value.GetType().IsPrimitive && value.GetType() != typeof(string))
                        {
                            var nested = SearchObjectForEngine(value, typeNames, depth + 1, maxDepth);
                            if (nested != null) return nested;
                        }
                    }
                    catch { /* skip inaccessible fields */ }
                }

                return null;
            }

            private static void DiscoverPictureBoxes(System.Windows.Forms.Form form)
            {
                try
                {
                    var pbs = new System.Collections.Generic.List<System.Windows.Forms.Control>();

                    // Strategy 1: look for a PictureBox[] or Control[] array field
                    var flags = System.Reflection.BindingFlags.Instance |
                                System.Reflection.BindingFlags.NonPublic |
                                System.Reflection.BindingFlags.Public;

                    foreach (var field in form.GetType().GetFields(flags))
                    {
                        try
                        {
                            var val = field.GetValue(form);
                            if (val is System.Windows.Forms.PictureBox[] pbArr)
                            {
                                _pictureBoxes = pbArr;
                                Log("[ENGINE] Found PictureBox[] array with " + pbArr.Length + " elements in field '" + field.Name + "'.");
                                return;
                            }
                        }
                        catch { }
                    }

                    // Strategy 2: collect individual PictureBox fields in declaration order
                    foreach (var field in form.GetType().GetFields(flags))
                    {
                        try
                        {
                            var val = field.GetValue(form);
                            if (val is System.Windows.Forms.PictureBox)
                                pbs.Add((System.Windows.Forms.Control)val);
                        }
                        catch { }
                    }

                    if (pbs.Count > 0)
                    {
                        _pictureBoxes = pbs.ToArray();
                        Log("[ENGINE] Collected " + pbs.Count + " PictureBox fields for SHOW/HIDE.");
                    }
                    else
                    {
                        Log("[ENGINE] No PictureBox controls found on form.");
                    }
                }
                catch (Exception ex)
                {
                    Log("[ENGINE] PictureBox discovery error: " + ex.Message);
                }
            }

            private static void StoreEngine(object engine)
            {
                _speechEngine = engine;

                // Resolve EmulateRecognize(string) — present on both engine types
                var method = engine.GetType().GetMethod(
                    "EmulateRecognize",
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public,
                    null,
                    new System.Type[] { typeof(string) },
                    null);

                if (method != null)
                {
                    _emulateMethod = method;
                    Log("[ENGINE] Ready: " + engine.GetType().FullName + ".EmulateRecognize(string)");
                }
                else
                {
                    Log("[ENGINE] Engine found but EmulateRecognize(string) not available on " + engine.GetType().FullName);
                }

                // Also locate EmulateRecognizeAsync(string) — this one works while the
                // engine is actively doing asynchronous recognition (the sync version throws).
                var asyncMethod = engine.GetType().GetMethod(
                    "EmulateRecognizeAsync",
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public,
                    null,
                    new System.Type[] { typeof(string) },
                    null);

                if (asyncMethod != null)
                {
                    _emulateAsyncMethod = asyncMethod;
                    Log("[ENGINE] Ready: " + engine.GetType().FullName + ".EmulateRecognizeAsync(string)");
                }
                else
                {
                    Log("[ENGINE] EmulateRecognizeAsync(string) not found on " + engine.GetType().FullName);
                }

                // Also locate RecognizeAsyncCancel() and RecognizeAsync(RecognizeMode)
                // for the stop/restart fallback strategy
                _recognizeAsyncCancelMethod = engine.GetType().GetMethod(
                    "RecognizeAsyncCancel",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public,
                    null, System.Type.EmptyTypes, null);
                if (_recognizeAsyncCancelMethod != null)
                    Log("[ENGINE] Ready: RecognizeAsyncCancel()");

                // RecognizeAsync(RecognizeMode) – RecognizeMode is an enum: Single=0, Multiple=1
                var recognizeMethods = engine.GetType().GetMethods(
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
                foreach (var rm in recognizeMethods)
                {
                    if (rm.Name == "RecognizeAsync")
                    {
                        var rmParams = rm.GetParameters();
                        if (rmParams.Length == 1 && rmParams[0].ParameterType.IsEnum)
                        {
                            _recognizeAsyncMethod = rm;
                            Log("[ENGINE] Ready: RecognizeAsync(" + rmParams[0].ParameterType.FullName + ")");
                            break;
                        }
                    }
                }

                // Log all public methods on engine for diagnostics
                Log("[ENGINE] Speech engine public methods:");
                foreach (var em in recognizeMethods)
                {
                    if (em.DeclaringType == engine.GetType() || em.DeclaringType.FullName.Contains("Speech"))
                        Log("[ENGINE]   " + em.Name + "(" + string.Join(", ", System.Array.ConvertAll(em.GetParameters(), p => p.ParameterType.Name)) + ") → " + em.ReturnType.Name);
                }
            }

            /// <summary>
            /// Executes an animation script file.
            /// Supported directives:
            ///   OPEN_URL &lt;url&gt;         - opens URL in default browser
            ///   PLAY_AUDIO &lt;rel-path&gt;  - plays a .wav file
            ///   WAIT &lt;ms&gt;              - sleeps
            ///   RUN &lt;path&gt;            - runs an executable/bat
            ///   HIDE_ALL / SHOW N / HIDE N  - UI animation (logged, not rendered here)
            /// </summary>
            private static void RunAnimationScript(string scriptPath)
            {
                var lines = File.ReadAllLines(scriptPath);
                foreach (var rawLine in lines)
                {
                    var line = rawLine.Trim();
                    if (line.Length == 0 || line.StartsWith("#")) continue;

                    var spaceIdx = line.IndexOf(' ');
                    var directive = spaceIdx >= 0 ? line.Substring(0, spaceIdx).ToUpperInvariant() : line.ToUpperInvariant();
                    var arg       = spaceIdx >= 0 ? line.Substring(spaceIdx + 1).Trim() : "";

                    try
                    {
                        switch (directive)
                        {
                            case "OPEN_URL":
                                Log("[OPEN_URL] " + arg);
                                System.Diagnostics.Process.Start(arg);
                                break;

                            case "PLAY_AUDIO":
                                var audioPath = System.IO.Path.Combine(BaseDir, arg.Replace('/', System.IO.Path.DirectorySeparatorChar));
                                Log("[PLAY_AUDIO] " + audioPath);
                                if (File.Exists(audioPath))
                                {
                                    var player = new System.Media.SoundPlayer(audioPath);
                                    player.Play();
                                }
                                else
                                    Log("[WARN] Audio file not found: " + audioPath);
                                break;

                            case "WAIT":
                                int ms;
                                if (int.TryParse(arg, out ms))
                                    System.Threading.Thread.Sleep(ms);
                                break;

                            case "RUN":
                                var runPath = System.IO.Path.Combine(BaseDir, arg.Replace('/', System.IO.Path.DirectorySeparatorChar));
                                Log("[RUN] " + runPath);
                                System.Diagnostics.Process.Start(runPath);
                                break;

                            case "HIDE_ALL":
                                if (_pictureBoxes != null && _mainForm != null)
                                {
                                    _mainForm.Invoke(new Action(() =>
                                    {
                                        foreach (var ctrl in _pictureBoxes)
                                            if (ctrl != null) ctrl.Visible = false;
                                    }));
                                }
                                break;

                            case "SHOW":
                            {
                                int showIdx;
                                if (int.TryParse(arg, out showIdx) && _pictureBoxes != null
                                    && showIdx >= 0 && showIdx < _pictureBoxes.Length && _mainForm != null)
                                {
                                    var target = _pictureBoxes[showIdx];
                                    _mainForm.Invoke(new Action(() => { if (target != null) target.Visible = true; }));
                                }
                                break;
                            }

                            case "HIDE":
                            {
                                int hideIdx;
                                if (int.TryParse(arg, out hideIdx) && _pictureBoxes != null
                                    && hideIdx >= 0 && hideIdx < _pictureBoxes.Length && _mainForm != null)
                                {
                                    var target = _pictureBoxes[hideIdx];
                                    _mainForm.Invoke(new Action(() => { if (target != null) target.Visible = false; }));
                                }
                                break;
                            }

                            default:
                                Log("[SCRIPT] Unknown directive: " + directive);
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("[ERROR] Script directive '" + directive + "' failed: " + ex.Message);
                    }
                }
            }

            // ── Helpers ───────────────────────────────────────────────────

            private static void EnsureInputFile()
            {
                try { System.IO.Directory.CreateDirectory(BaseDir); } catch { }

                if (!File.Exists(InputFile))
                    File.WriteAllText(InputFile, "");

                if (!File.Exists(CommandsFile))
                    Log("[WARN] custom-commands/commands.txt not found – no commands will be matched.");
            }

            private static void Log(string msg)
            {
                var line = $"[{DateTime.Now:HH:mm:ss.fff}] {msg}";
                Console.WriteLine(line);
                try { File.AppendAllText(LogFile, line + "\n"); }
                catch { /* ignore log-write failures */ }
            }
        }
        """;
}
