Unit tests should never pop a `MsgBox` or prompt with an `InputBox`, but what if you're trying to write a test for a method that _does_ make these calls?

You could write yourself an `IMsgBox` abstraction and tests could inject an implementation that doesn't actually call `MsgBox`, but then you also need to implement and inject another class that does.

By using the `Rubberduck.FakesProvider` API in your unit tests, you can ignore actual `MsgBox` calls and even configure their output to suit your needs, effectively removing the need for an abstraction here. The "fake" message box also tracks invocations, so you can assert that a message was shown with the expected parameterization.

There are many more calls that can be intercepted by Rubberduck to give greater control over unit testing. See the tables below for more details.

### How?!

When this API in use, Rubberduck uses the EasyHook library to hook into the VBE DLL and intercepts internal calls to functions like `rtcMsgBox`, and overrides them with the behavior configured by a Rubberduck unit test.

No hooks are active when there isn't a test running, or when a test doesn't use the Fakes API.

### Setting up and verifying a fake

The standard test module template defines Assert and Fakes private fields. When changed to be early-bound (needs a reference to the Rubberduck type library), the declarations and initialization look like this:

```
'@TestModule
Option Explicit
Option Private Module
Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider
'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub
```

An example sub that we might want to test is `SayHello` which displays an `InputBox` and then displays a `MsgBox`. We'd like to test the output message automatically...

```
Public Sub SayHello()
    Dim name As String
    name = InputBox("Enter your name")
    MsgBox "Hello " + name
End Sub
```

Out test method needs to first setup the behaviour of these VBA functions before calling the `SayHello` sub. Then we can even check the output is as expected. We add this code to the test module:

```
'@TestMethod
Private Sub TestSayHello()
    'Setup the behaviour of methods we want to fake
    Fakes.MsgBox.Returns vbYes
    Fakes.InputBox.Returns "Rubberduck Fan"
    
    'Call the sub we want to test
    SayHello
    
    'Check the output
    Fakes.MsgBox.Verify.Parameter Fakes.Params.MsgBox.Prompt, "Hello Rubberduck Fan"
End Sub
```

When executing this method from the Rubberduck Test Explorer window, we will see that the test passes and there is no user interaction at all!

### What else can we fake?

Name | Description | Parameter names
--- | --- | ---
MsgBox | Configures VBA.Interaction.MsgBox calls | Fakes.Params.MsgBox
InputBox | Configures VBA.Interaction.InputBox calls | Fakes.Params.InputBox
Beep | Configures VBA.Interaction.Beep calls | 
Environ | Configures VBA.Interaction.Environ calls | Fakes.Params.Environ
Timer | Configures VBA.DateTime.Timer calls | 
DoEvents | Configures VBA.Interaction.DoEvents calls | 
Shell | Configures VBA.Interaction.Shell calls | Fakes.Params.Shell
SendKeys | Configures VBA.Interaction.SendKeys calls | Fakes.Params.SendKeys
Kill | Configures VBA.FileSystem.Kill calls | Fakes.Params.Kill
MkDir | Configures VBA.FileSystem.MkDir calls | Fakes.Params.MkDir
RmDir | Configures VBA.FileSystem.RmDir calls | Fakes.Params.RmDir
ChDir | Configures VBA.FileSystem.ChDir calls | Fakes.Params.ChDir
ChDrive | Configures VBA.FileSystem.ChDrive calls | Fakes.Params.ChDrive
CurDir | Configures VBA.FileSystem.CurDir calls | Fakes.Params.CurDir
Now | Configures VBA.DateTime.Now calls | 
Time | Configures VBA.DateTime.Time calls | 
Date | Configures VBA.DateTime.Date calls | 
Rnd* | Configures VBA.Math.Rnd calls | Fakes.Params.Rnd
DeleteSetting* | Configures VBA.Interaction.DeleteSetting calls | Fakes.Params.DeleteSetting
SaveSetting* | Configures VBA.Interaction.SaveSetting calls | Fakes.Params.SaveSetting
Randomize* | Configures VBA.Math.Randomize calls | Fakes.Params.Randomize
GetAllSettings* | Configures VBA.Interaction.GetAllSettings calls | 
SetAttr* | Configures VBA.FileSystem.SetAttr calls | Fakes.Params.SetAttr
GetAttr* | Configures VBA.FileSystem.GetAttr calls | Fakes.Params.GetAttr
FileLen* | Configures VBA.FileSystem.FileLen calls | Fakes.Params.FileLen
FileDateTime* | Configures VBA.FileSystem.FileDateTime calls | Fakes.Params.FileDateTime
FreeFile* | Configures VBA.FileSystem.FreeFile calls | Fakes.Params.FreeFile
IMEStatus* | Configures VBA.Information.IMEStatus calls | 
Dir* | Configures VBA.FileSystem.Dir calls | Fakes.Params.Dir
FileCopy* | Configures VBA.FileSystem.FileCopy calls | Fakes.Params.FileCopy

Note: Items marked with * are currently only available in pre-release builds. Similarly for the `Fakes.Params` parameter names. 

### Verification methods

Sometimes you will just be checking the result of the method being tested in which case the Rubberduck Assert class provides various test methods. Further test functionality is provided through the `Verify` interface which tracks how often a faked method is called and what the parameters passed where e.g. `Fakes.Beep.Verify.Once` would be a test that the `Beep` function was called exactly once. The full set of verification methods available through the `Verify` interface is:

Name | Description
--- | ---
AtLeast | Verifies that the faked procedure was called a specified minimum number of times.
AtLeastOnce | Verifies that the faked procedure was called one or more times.
AtMost | Verifies that the faked procedure was called a specified maximum number of times.
AtMostOnce | Verifies that the faked procedure was not called or was only called once.
Between | Verifies that the number of times the faked procedure was called falls within the supplied range.
Exactly | Verifies that the faked procedure was called a specified number of times.
Never | Verifies that the faked procedure was called exactly 0 times.
Once | Verifies that the faked procedure was called exactly one time.
Parameter | Verifies that the value of a given parameter to the faked procedure matches a specific value.
ParameterInRange | Verifies that the value of a given parameter to the faked procedure falls within a specified range.
ParameterIsPassed | Verifies that an optional parameter was passed to the faked procedure. The value is not evaluated.
ParameterIsType | Verifies that the passed value of a given parameter was of a type that matches the given type name.

The `Parameter`, `ParameterInRange`, `ParameterIsPassed`, and `ParameterIsType` methods take the name of the parameter to check as a string. To make this easy to get right, the `Fakes.Params` property can be used to provide these strings with the benefit of Intellisense.

### Additional features

The IFake interface exposes members for the setup/configuration of fakes:

Name | Description
--- | ---
AssignsByRef | *Not implemented yet.* Configures the fake such as an invocation assigns the specified value to the specified ByRef argument.
Passthrough | Gets/sets whether invocations should pass through to the native call.
RaisesError | Configures the fake such as an invocation raises the specified run-time error.
Returns | Configures the fake such as the specified invocation returns the specified value.
ReturnsWhen | Configures the fake such as the specified invocation returns the specified value given a specific parameter value.
Verify | Gets an interface for verifying invocations performed during the test. See Verification methods above.

The `Returns` and `ReturnsWhen` methods are only available when the function being faked actually has a return value.

If you have repeated calls to a VBA function in the method to be tested, it is possible to set the return values specifically for each call. For example, if your code calls the `Rnd` function twice to control two nested branches within your code, you can check all possible paths through the code by setting different return values for specific invocations of the VBA function through the optional `Invocation` parameter e.g. `Fakes.Rnd.Returns 0.1, 1` and `Fakes.Rnd.Returns 0.9, 2` would set separate return values for the first and second calls respectively.