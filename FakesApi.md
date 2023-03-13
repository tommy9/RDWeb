Unit tests should never pop a `MsgBox` or prompt with an `InputBox`, but what if you're trying to write a test for a method that _does_ make these calls?

You could write yourself an `IMsgBox` abstraction and tests could inject an implementation that doesn't actually call `MsgBox`, but then you also need to implement and inject another class that does.

By using the `Rubberduck.FakesProvider` API in your unit tests, you can ignore actual `MsgBox` calls and even configure their output to suit your needs, effectively removing the need for an abstraction here. The "fake" message box also tracks invocations, so you can assert that a message was shown with the expected parameterization.

There are many more calls that can be intercepted by Rubberduck to give greater control over unit testing. See the tables below for more details.

### How?!

When this API in use, Rubberduck uses the EasyHook library to hook into the VBE DLL and intercepts internal calls to functions like `rtcMsgBox`, and overrides them with the behavior configured by a Rubberduck unit test.

No hooks are active when there isn't a test running, or when a test doesn't use the Fakes API.

### Setting up a fake

The standard test module template defines Assert and Fakes private fields. When early-bound (needs a reference to the Rubberduck type library), the declarations and initialization look like this:

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

An example function that we might want to test is:

```
Public Sub SayHello()
    Dim name As String
End Sub
```

### Verifying calls to fakes

