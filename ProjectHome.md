# Introduction #
[AutoIt](http://www.autoitscript.com/autoit3/index.shtml) is a very useful automation scripting language for Microsoft Windows. It allows for GUI automation using a very simple syntax and can be useful for testing Windows applications. It is packaged with AutoItX which supports accessing AutoIt functions through COM objects.

AutoItX4Java uses [JACOB](http://sourceforge.net/projects/jacob-project/) to access AutoItX through COM and strives to provide a native Java interface while maintaining the simplicity of AutoIt. Getting started is simple.

  1. Download [JACOB](http://sourceforge.net/projects/jacob-project/).
  1. Download and install [AutoIt](http://www.autoitscript.com/autoit3/index.shtml).
  1. Add jacob.jar and autoitx4java.jar to your library path.
  1. Place the jacob-1.15-M4-x64.dll file in your library path.
  1. Start using AutoItX.

# Example #
```
        File file = new File("lib", "jacob-1.15-M4-x64.dll"); //path to the jacob dll
        System.setProperty(LibraryLoader.JACOB_DLL_PATH, file.getAbsolutePath());

        AutoItX x = new AutoItX();
        String notepad = "Untitled - Notepad";
        String testString = "this is a test.";
        x.run("notepad.exe");
        x.winActivate(notepad);
        x.winWaitActive(notepad);
        x.send(testString);
        Assert.assertTrue(x.winExists(notepad, testString));
        x.winClose(notepad, testString);
        x.winWaitActive("Notepad");
        x.send("{ALT}n");
        Assert.assertFalse(x.winExists(notepad, testString));
```

**Troubleshooting**

Both AutoItX3 and Jacob dll's come in x86 and x64 versions. Ensure you use the correct dlls.

If AutoItX isn't disposing itself properly make a call to ComThread.Release() when you are done using AutoItX.
Refer to JACOB documentation: [JacobThreading](http://danadler.com/jacob/JacobThreading.html) and [Object Lifetime](http://danadler.com/jacob/JacobComLifetime.html)

**Note**

If you do not want to install AutoIt, you can just grab AutoItX3.dll and register it with: regsvr32.exe AutoItX3.dll

AutoItX4Java is comprised of one file so you can either use the jar or copy the .java file into your source.

If you get errors like: "Can’t co-create object" then ensure you have the AuotItX dll registered. If you need to manually register the 64 bit version, use: "\Windows\SysWOW64\regsvr32.exe AutoItX3\_x64.dll" otherwise use "regsrv32.exe AutoItX3.dll".

**Warning**

AutoItX4Java is not completely tested. Use at your own risk.