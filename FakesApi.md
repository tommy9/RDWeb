Unit tests should never pop a `MsgBox` or prompt with an `InputBox`, but what if you're trying to write a test for a method that _does_ make these calls?

You could write yourself an `IMsgBox` abstraction and tests could inject an implementation that doesn't actually call `MsgBox`, but then you also need to implement and inject another class that does.

By using the `Rubberduck.FakesProvider` API in your unit tests, you can ignore actual `MsgBox` calls and even configure their output to suit your needs, effectively removing the need for an abstraction here. The "fake" message box also tracks invocations, so you can assert that a message was shown with the expected parameterization.

### How?!

When this API in use, Rubberduck uses the EasyHook library to hook into the VBE DLL and intercepts internal calls to functions like `rtcMsgBox`, and overrides them with the behavior configured by a Rubberduck unit test.

No hooks are active when there isn't a test running, or when a test doesn't use the Fakes API.