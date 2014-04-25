/**
 * Copyright (c) 2011 Richard Kanavati
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */
package autoitx4java;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.SafeArray;
import com.jacob.com.Variant;

/**
 * A Java AutoItX3 Bridge.
 * @author Richard Kanavati
 */
public class AutoItX {

    /**
     * The AutoItX Component.
     */
    protected ActiveXComponent autoItX;
    
    /**
     * Maximizes the specified window.
     */    
    public final static int SW_MAXIMIZE = 3;
    
    /**
     * Minimizes the specified window and activates the next top-level window 
     * in the Z order.
     */    
    public final static int SW_HIDE = 0;
    /**
     * Activates and displays the window. If the window is minimized or maximized, 
     * the system restores it to its original size and position. 
     * An application should specify this flag when restoring a minimized window.
     */    
    public final static int SW_RESTORE = 9;
    
    /**
     * Activates the window and displays it in its current size and position.
     */    
    public final static int SW_SHOW = 5;
    
    /**
     * Sets the show state based on the SW_ value specified by the program 
     * that started the application.
     */    
    public final static int SW_SHOWDEFAULT = 10;
    
    /**
     * Activates the window and displays it as a maximized window.
     */    
    public final static int SW_SHOWMAXIMIZED = 3;
    
    /**
     * Activates the window and displays it as a minimized window.
     */    
    public final static int SW_SHOWMINIMIZED = 2;
    
    /**
     * Displays the window as a minimized window. This value is similar to 
     * SW_SHOWMINIMIZED, except the window is not activated.
     */
    public final static int SW_SHOWMINNOACTIVE = 7;
    
    /**
     * Displays the window in its current size and position. This value is 
     * similar to SW_SHOW, except the window is not activated.
     */
    public final static int SW_SHOWNA = 8;
    
    /**
     * Displays a window in its most recent size and position. This value is 
     * similar to SW_SHOWNORMAL, except the window is not actived.
     */
    public final static int SW_SHOWNOACTIVATE = 4;
    
    /**
     * Activates and displays a window. If the window is minimized or maximized, 
     * the system restores it to its original size and position. 
     * An application should specify this flag when displaying the window 
     * for the first time.
     */
    public final static int SW_SHOWNORMAL = 1;
    
    /**
     * Sets the way coords are used in the caret functions, either absolute 
     * coords or coords relative to the current active window:
     * <br/>0 = relative coords to the active window,
     * <br/>1 = absolute screen coordinates (default),
     * <br/>2 = relative coords to the client area of the active window.
     */
    public final static String OPT_CARET_COORD_MODE = "CaretCoordMode";
    
    /**
     * Alters the length of the brief pause in between mouse clicks.
     * Time in milliseconds to pause (default=10).
     */
    public final static String OPT_MOUSE_CLICK_DELAY = "MouseClickDelay";
    
    /**
     * Alters the length of the brief pause in between mouse clicks.
     * Time in milliseconds to pause (default=10).
     */
    public final static String OPT_MOUSE_CLICK_DOWN_DELAY = "MouseClickDownDelay";
    
    /**
     * Alters the length of the brief pause at the start and end of a mouse 
     * drag operation. Time in milliseconds to pause (default=250).
     */
    public final static String OPT_MOUSE_CLICK_DRAG_DELAY = "MouseClickDragDelay";
    
    /**
     * Sets the way coords are used in the mouse functions, either absolute 
     * coords or coords relative to the current active window:
     * <br/>0 = relative coords to the active window
     * <br/>1 = absolute screen coordinates (default)
     * <br/>2 = relative coords to the client area of the active window
     */
    public final static String OPT_MOUSE_COORD_MODE = "MouseCoordMode";
    
    /**
     * Sets the way coords are used in the mouse functions, either absolute 
     * coords or coords relative to the current active window:
     * <br/>0 = relative coords to the active window
     * <br/>1 = absolute screen coordinates (default)
     * <br/>2 = relative coords to the client area of the active window
     */
    public final static String OPT_PIXEL_COORD_MODE = "PixelCoordMode";
    
    /**
     * Sets the way coords are used in the pixel functions, either absolute 
     * coords or coords relative to the current active window:
     * <br/>0 = relative coords to the active window
     * <br/>1 = absolute screen coordinates (default)
     * <br/>2 = relative coords to the client area of the active window
     */
    public final static String OPT_SEND_ATTACH_MODE = "SendAttachMode";
    
    /**
     * Specifies if AutoIt attaches input threads when using then Send() 
     * function. When not attaching (default mode=0) detecting the state of 
     * capslock/scrolllock and numlock can be unreliable under NT4. However, 
     * when you specify attach mode=1 the Send("{... down/up}") syntax will not 
     * work and there may be problems with sending keys to "hung" windows. 
     * ControlSend() ALWAYS attaches and is not affected by this mode.
     * <br/>0 = don't attach (default)
     * <br/>1 = attach
     */
    public final static String OPT_SEND_CAPSLOCK_MODE = "SendCapslockMode";
    
    /**
     * Specifies if AutoIt should store the state of capslock before a 
     * Send function and restore it afterwards.
     * <br/>0 = don't store/restore
     * <br/>1 = store and restore (default)
     */
    public final static String OPT_SEND_KEY_DELAY = "SendKeyDelay";
    
    /**
     * Alters the length of time a key is held down before released during a 
     * keystroke. For applications that take a while to register keypresses 
     * (and many games) you may need to raise this value from the default.
     * Time in milliseconds to pause (default=1).
     */
    public final static String OPT_SEND_KEY_DOWN_DELAY = "SendKeyDownDelay";
    
    /**
     * Specifies if hidden window text can be "seen" by the window matching functions.
     * <br/>0 = Do not detect hidden text (default)
     * <br/>1 = Detect hidden text
     */
    public final static String OPT_WIN_DETECT_HIDDEN_TEXT = "WinDetectHiddenText";
    
    /**
     * Allows the window search routines to search child windows as well as top-level windows.
     * <br/>0 = Only search top-level windows (default)
     * <br/>1 = Search top-level and child windows
     */
    public final static String OPT_WIN_SEARCH_CHILDREN = "WinSearchChildren";
    
    /**
     * Alters the method that is used to match window text during search operations.
     * <br/>1 = Complete / Slow mode (default)
     * <br/>2 = Quick mode
     * In quick mode AutoIt can usually only "see" dialog text, button text 
     * and the captions of some controls. In the default mode much more text 
     * can be seen (for instance the contents of the Notepad window).
     * If you are having performance problems when performing many window 
     * searches then changing to the "quick" mode may help.
     */
    public final static String OPT_WIN_TEXT_MATCH_MODE = "WinTextMatchMode";
    
    /**
     * Alters the method that is used to match window titles during search operations.
     * <br/>1 = Match the title from the start (default)
     * <br/>2 = Match any substring in the title
     * <br/>3 = Exact title match
     * <br/>4 = Advanced mode, see Window Titles & Text (Advanced)
     */
    public final static String OPT_WIN_TITLE_MATCH_MODE = "WinTitleMatchMode";
    
    /**
     * Alters how long a script should briefly pause after a successful window-related operation.
     * Time in milliseconds to pause (default=250).
     */
    public final static String OPT_WIN_WAIT_DELAY = "WinWaitDelay";
   

    /**
     * Initializes the AutoItX Jacob COM object.
     */
    public AutoItX() {
        autoItX = new ActiveXComponent("AutoItX3.Control");
    }
    
    /**
     * The AutoItX version.
     * @return The AutoItX version.
     */
    public String getVersion() {
        return autoItX.getProperty("version").getString();
    }

    /**
     * Value of the error flag.
     * @return The value of the error flag.
     */
    public int getError() {
        Variant error = autoItX.invoke("error");
        return error.getInt();
    }

    /**
     * Retrieves text from the clipboard.
     * @return A string containing the text on the clipboard. Sets error to 1 
     * if clipboard is empty or contains a non-text entry.
     */
    public String clipGet() {
       return autoItX.invoke("ClipGet").getString();
    }

    /**
     * Writes text to the clipboard.
     * @param value The text to write to the clipboard.
     */
    public void clipPut(String value) {
        autoItX.invoke("ClipPut", new Variant(value));
    }

    /**
     * Maps a network drive.
     * <p><b>Doesn't always work</b></p>
     * @param device The device to map, for example "O:" or "LPT1:". If you 
     *               pass a blank string for this parameter a connection is made but not 
     *               mapped to a specific drive. If you specify "*" an unused drive letter 
     *               will be automatically selected.
     * @param remote The remote share to connect to in the form "\\server\share".
     * @return True if success, false otherwise
     */
    public boolean driveMapAdd(String device, String remote) {
        return driveMapAdd(device, remote, 0, "", "");
    }

    /**
     * Maps a network drive.     
     * @param device The device to map, for example "O:" or "LPT1:". 
     *               If you pass a blank string for this parameter a connection 
     *               is made but not mapped to a specific drive. 
     *               If you specify "*" an unused drive letter will be automatically 
     *               selected.
     * @param remote The remote share to connect to in the form "\\server\share".
     * @param flags A combination of the following: 0 = default, 
     *              1 = Persistent mapping, 
     *              8 = Show authentication dialog if required
     * @param username The username
     * @param password The password
     * @return True if success, false otherwise
     */
    public boolean driveMapAdd(String device, String remote, int flags, String username, String password) {
        Variant vDevice = new Variant(device);
        Variant vRemote = new Variant(remote);
        Variant vFlags = new Variant(flags);
        Variant vUsername = new Variant(username);
        Variant vPassword = new Variant(password);
        Variant[] params = new Variant[]{vDevice, vRemote, vFlags, vUsername, vPassword};
        Variant result = autoItX.invoke("DriveMapAdd", params);
        if (result.getvt() == Variant.VariantString) {
            return oneToTrue(Integer.parseInt(result.getString()));
        }
        
        return oneToTrue(result.getInt());
    }

    /**
     * Maps a network drive.
     * @param device The device to map, for example "O:" or "LPT1:". 
     *               If you pass a blank string for this parameter a connection 
     *               is made but not mapped to a specific drive. If you specify 
     *               "*" an unused drive letter will be automatically selected.
     * @param remote The remote share to connect to in the form "\\server\share".
     * @param flags A combination of the following:0 = default, 
     *              1 = Persistent mapping, 8 = Show authentication dialog if required
     * @return True if success, false otherwise
     */
    public boolean driveMapAdd(String device, String remote, int flags) {
        Variant vDevice = new Variant(device);
        Variant vRemote = new Variant(remote);
        Variant vFlags = new Variant(flags);
        Variant[] params = new Variant[]{vDevice, vRemote, vFlags};
        Variant result = autoItX.invoke("DriveMapAdd", params);
        if (result.getvt() == Variant.VariantString) {
            return oneToTrue(Integer.parseInt(result.getString()));
        }
        
        return oneToTrue(result.getInt());
    }

    /**
     * Disconnects a network drive.
     * @param device The device to disconnect, e.g. "O:" or "LPT1:".
     * @return True if success, false otherwise
     */
    public boolean driveMapDelete(String device) {
        Variant result = autoItX.invoke("DriveMapDel", device);
        return oneToTrue(result);
    }

    /**
     * The device (drive or printer) letter to query. Eg. "O:" or "LPT1:"
     * @param device The device to disconnect, e.g. "O:" or "LPT1:".
     * @return Details of the mapping, e.g. \\server\share. If blank, sets .error() to 1.
     */
    public String driveMapGet(String device) {
        Variant result = autoItX.invoke("DriveMapGet", device);
        return result.getString();
    }

    /**
     * Deletes a value from a standard format .ini file.
     * @param filename The filename of the .ini file.
     * @param section The section name in the .ini file.
     * @param key The key name in the in the .ini file. 
     *            If no key name is given the entire section is deleted.
     * @return True if success, false otherwise
     */
    public boolean iniDelete(String filename, String section, String key) {
        Variant vFilename = new Variant(filename);
        Variant vSection = new Variant(section);
        Variant vKey = new Variant(key);
        Variant[] params = new Variant[]{vFilename, vSection, vKey};
        Variant result = autoItX.invoke("IniDelete", params);
        return oneToTrue(result);
    }

    /**
     * Deletes a value from a standard format .ini file.
     * @param filename The filename of the .ini file.
     * @param section The section name in the .ini file.
     * @return True if success, false otherwise
     */
    public boolean iniDelete(String filename, String section) {
        return iniDelete(filename, section, "");
    }

    /**
     * Reads a value from a standard format .ini file.
     * @param filename The filename of the .ini file.
     * @param section The section name in the .ini file.
     * @param key The key name in the in the .ini file.
     * @param defaultVal The default value to return if the requested key is not found.
     * @return The requested key value if found, otherwise the default value.
     */
    public String iniRead(String filename, String section, String key, String defaultVal) {
        Variant vFilename = new Variant(filename);
        Variant vSection = new Variant(section);
        Variant vKey = new Variant(key);
        Variant vDefault = new Variant(defaultVal);
        Variant[] params = new Variant[]{vFilename, vSection, vKey, vDefault};
        Variant result = autoItX.invoke("IniRead", params);
        return result.getString();
    }

    /**
     * Writes a value to a standard format .ini file.
     * @param filename The filename of the .ini file.
     * @param section The section name in the .ini file.
     * @param key The key name in the in the .ini file.
     * @param value The value to write/change.
     * @return True if success, false otherwise.
     */
    public Boolean iniWrite(String filename, String section, String key, String value) {
        Variant vFilename = new Variant(filename);
        Variant vSection = new Variant(section);
        Variant vKey = new Variant(key);
        Variant vValue = new Variant(value);
        Variant[] params = new Variant[]{vFilename, vSection, vKey, vValue};
        Variant result = autoItX.invoke("IniWrite", params);
        return oneToTrue(result);
    }

    /**
     * Generates a checksum for a region of pixels.
     * @param left Left coordinate of rectangle.
     * @param top Top coordinate of rectangle.
     * @param right Right coordinate of rectangle.
     * @param bottom Bottom coordinate of rectangle.
     * @param step Instead of checksumming each pixel use a value larger than 
     *             1 to skip pixels (for speed). E.g. A value of 2 will only 
     *             check every other pixel. Default is 1.
     * @return The checksum value of the region.
     */
    public double pixelChecksum(int left, int top, int right, int bottom, int step) {
        Variant vLeft = new Variant(left);
        Variant vTop = new Variant(top);
        Variant vRight = new Variant(right);
        Variant vBottom = new Variant(bottom);
        Variant vStep = new Variant(step);
        Variant[] params = new Variant[]{vLeft, vTop, vRight, vBottom, vStep};
        Variant result = autoItX.invoke("PixelChecksum", params);
        return result.getDouble();
    }

    /**
     * Generates a checksum for a region of pixels.
     * @param left Left coordinate of rectangle.
     * @param top Top coordinate of rectangle.
     * @param right Right coordinate of rectangle.
     * @param bottom Bottom coordinate of rectangle.
     * @return The checksum value of the region.
     */
    public double pixelChecksum(int left, int top, int right, int bottom) {
        return pixelChecksum(left, top, right, bottom, 0);
    }

    /**
     * Returns a pixel color according to x,y pixel coordinates.
     * @param x x coordinate of pixel.
     * @param y y coordinate of pixel.
     * @return Decimal value of pixel's color.
     */
    public float pixelGetColor(int x, int y) {
        Variant vX = new Variant(x);
        Variant vY = new Variant(y);
        Variant[] params = new Variant[]{vX, vY};
        Variant result = autoItX.invoke("PixelGetColor", params);
        return result.getInt();
    }

    /**
     * Searches a rectangle of pixels for the pixel color provided.
     * @param left left coordinate of rectangle.
     * @param top top coordinate of rectangle.
     * @param right right coordinate of rectangle.
     * @param bottom bottom coordinate of rectangle.
     * @param color Color value of pixel to find (in decimal or hex).
     * @param shadeVariation A number between 0 and 255 to indicate the allowed 
     *                       number of shades of variation of the red, green, 
     *                       and blue components of the colour. 
     *                       Default is 0 (exact match).
     * @param step Instead of searching each pixel use a value larger than 1 
     *             to skip pixels (for speed). E.g. A value of 2 will only 
     *             check every other pixel. Default is 1.
     * @return The pixel's coordinates in a 2 element array, otherwise sets 
     * .error() to one.
     */
    public long[] pixelSearch(int left, int top, int right, int bottom, 
            int color, int shadeVariation, int step) {
        Variant vLeft = new Variant(left);
        Variant vTop = new Variant(top);
        Variant vRight = new Variant(right);
        Variant vBottom = new Variant(bottom);
        Variant vColor = new Variant(color);
        Variant vShadeVariation = new Variant(shadeVariation);
        Variant vStep = new Variant(step);
        Variant[] params = new Variant[]{vLeft, vTop, vRight, vBottom, vColor, vShadeVariation, vStep};
        Variant result = autoItX.invoke("PixelSearch", params);
        long[] l = new long[2];
        if (result.getvt() == 8204) {
            l[0] = result.toSafeArray().getLong(0);
            l[1] = result.toSafeArray().getLong(1);
        }
        
        return l;
    }

    /**
     * Searches a rectangle of pixels for the pixel color provided.
     * @param left left coordinate of rectangle.
     * @param top top coordinate of rectangle.
     * @param right right coordinate of rectangle.
     * @param bottom bottom coordinate of rectangle.
     * @param color Color value of pixel to find (in decimal or hex).
     * @return The pixel's coordinates in a 2 element array, otherwise 
     * sets .error() to one.
     */
    public long[] pixelSearch(int left, int top, int right, int bottom, int color) {
        return pixelSearch(left, top, right, bottom, color, 0, 1);
    }

    /**
     * Sends simulated keystrokes to the active window.
     * @param keys The sequence of keys to send.
     * @param isRaw Changes how "keys" is processed: 
     * true -  Text contains special characters like + and ! to indicate SHIFT 
     * and ALT key presses. false, keys are sent raw.
     */
    public void send(String keys, boolean isRaw) {
        Variant vKeys = new Variant(keys);
        Variant vFlag = new Variant(isRaw ? 1 : 0);
        Variant[] params = new Variant[]{vKeys, vFlag};
        autoItX.invoke("Send", params);
    }

    /**
     * Sends simulated keystrokes to the active window.
     * @param keys The sequence of keys to send.
     */
    public void send(String keys) {
        send(keys, true);
    }

    /**
     * Creates a tooltip anywhere on the screen.
     * @param text The text of the tooltip. (An empty string clears a displaying tooltip)
     * @param x The x,y position of the tooltip.
     * @param y The x,y position of the tooltip.
     */
    public void toolTip(String text, int x, int y) {
        Variant vText = new Variant(text);
        Variant vX = new Variant(x);
        Variant vY = new Variant(y);
        Variant[] params = new Variant[]{vText, vX, vY};
        autoItX.invoke("ToolTip", params);
    }

    /**
     * Creates a tooltip anywhere on the screen.
     * @param text The text of the tooltip. (An empty string clears a displaying tooltip)
     */
    public void toolTip(String text) {
        autoItX.invoke("ToolTip", text);
    }

    /**
     * Disable/enable the mouse and keyboard.
     * Requires admin rights in Vista and Windows 7.
     * @param disableInput True = disable user input, false = re-enable it.
     */
    public void blockInput(boolean disableInput) {
        autoItX.invoke("BlockInput", (disableInput ? 1 : 0));
    }

    /**
     * Opens or closes the CD tray.
     * @param drive The drive letter of the CD tray to control, in the format D:, E:, etc.
     * @param status Specifies if you want the CD tray to be open or closed: "open" or "closed"
     * @return True if success, false if drive is locked via CD burning 
     * software or if the drive letter is not a CD drive.
     */
    public boolean cdTray(String drive, String status) {
        Variant vDrive = new Variant(drive);
        Variant vStatus = new Variant(status);
        Variant[] params = new Variant[]{vDrive, vStatus};
        Variant result = autoItX.invoke("CDTray", params);
        return oneToTrue(result);
    }

    /**
     * Checks if the current user has administrator privileges.
     * @return True if is admin, false otherwise.
     */
    public boolean isAdmin() {
        return oneToTrue(autoItX.invoke("IsAdmin"));
    }

    /**
     * Changes the operation of various AutoIt functions/parameters.
     * @param option The option to change.
     * @param param The parameter (varies by option).
     * @return Value of the previous setting.
     */
    public String autoItSetOption(String option, String param) {
        Variant vOption = new Variant(option);
        Variant vParam = new Variant(param);
        Variant[] params = new Variant[]{vOption, vParam};
        Variant result = autoItX.invoke("AutoItSetOption", params);
        if (result.getvt() == Variant.VariantInt) {
            return String.valueOf(result.getInt());
        }
        
        return result.getString();
    }

    /**
     * Changes the operation of various AutoIt functions/parameters.
     * @param option The option to change.
     * @param param The parameter (varies by option).
     * @return Value of the previous setting.
     */
    public String setOption(String option, String param) {
        return autoItSetOption(option, param);
    }

    /**
     * Perform a mouse click operation.
     * @param button The button to click: "left", "right", "middle", "main", 
     *               "menu", "primary", "secondary".
     * @param x The x/y coordinates to move the mouse to. 
     *          If no x and y coords are given, the current position is used.
     * @param y The x/y coordinates to move the mouse to. 
     *          If no x and y coords are given, the current position is used.
     * @param clicks The number of times to click the mouse. Default is 1.
     * @param speed The speed to move the mouse in the range 1 (fastest) to 100 (slowest). 
     * A speed of 0 will move the mouse instantly. Default speed is 10.
     */
    public void mouseClick(String button, int x, int y, int clicks, int speed) {
        Variant vButton = new Variant(button);
        Variant vX = new Variant(x);
        Variant vY = new Variant(y);
        Variant vClicks = new Variant(clicks);
        Variant vSpeed = new Variant(speed);
        Variant[] params = new Variant[]{vButton, vX, vY, vClicks, vSpeed};
        autoItX.invoke("MouseClick", params);
    }

    /**
     * Perform a mouse click operation.
     * @param button The button to click: "left", "right", "middle", "main", 
     *               "menu", "primary", "secondary".
     * @param clicks The number of times to click the mouse. Default is 1.
     * @param speed The speed to move the mouse in the range 1 (fastest) to 100 (slowest). 
     * A speed of 0 will move the mouse instantly. Default speed is 10.
     */
    public void mouseClick(String button, int clicks, int speed) {
        Variant vButton = new Variant(button);
        Variant vClicks = new Variant(clicks);
        Variant vSpeed = new Variant(speed);
        Variant[] params = new Variant[]{vButton, vClicks, vSpeed};
        autoItX.invoke("MouseClick", params);
    }

    /**
     * Perform a mouse click and drag operation.
     * @param button The button to click: "left", "right", "middle", 
     *               "main", "menu", "primary", "secondary".
     * @param x The x/y coords to start the drag operation from.
     * @param y The x/y coords to start the drag operation from.
     * @param x2 The x/y coords to start the drag operation to.
     * @param y2 The x/y coords to start the drag operation to.
     * @param speed The speed to move the mouse in the range 1 (fastest) to 
     * 100 (slowest). A speed of 0 will move the mouse instantly. Default speed is 10.
     */
    public void mouseClickDrag(String button, int x, int y, int x2, int y2, int speed) {
        Variant vButton = new Variant(button);
        Variant vX = new Variant(x);
        Variant vY = new Variant(y);
        Variant vX2 = new Variant(x2);
        Variant vY2 = new Variant(y2);
        Variant vSpeed = new Variant(speed);
        Variant[] params = new Variant[]{vButton, vX, vY, vX2, vY2, vSpeed};
        autoItX.invoke("MouseClickDrag", params);
    }

    /**
     * Perform a mouse click and drag operation.
     * @param button The button to click: "left", "right", "middle", "main", 
     *               "menu", "primary", "secondary".
     * @param x The x/y coords to start the drag operation from.
     * @param y The x/y coords to start the drag operation from.
     * @param x2 The x/y coords to start the drag operation to.
     * @param y2 The x/y coords to start the drag operation to.
     */
    public void mouseClickDrag(String button, int x, int y, int x2, int y2) {
        mouseClickDrag(button, x, y, x2, y2, 10);
    }

    /**
     * Perform a mouse down event at the current mouse position.
     * @param button The button to click: "left", "right", "middle", "main", 
     *               "menu", "primary", "secondary".
     */
    public void mouseDown(String button) {
        autoItX.invoke("MouseDown", button);
    }

    /**
     * Returns a cursor ID Number of the current Mouse Cursor.
     * @return  0 = UNKNOWN (this includes pointing and grabbing hand icons)
     * <br/> 1 = APPSTARTING
     * <br/> 2 = ARROW
     * <br/> 3 = CROSS
     * <br/> 4 = HELP
     * <br/> 5 = IBEAM
     * <br/> 6 = ICON
     * <br/> 7 = NO
     * <br/> 8 = SIZE
     * <br/> 9 = SIZEALL
     * <br/> 10 = SIZENESW
     * <br/> 11 = SIZENS
     * <br/> 12 = SIZENWSE
     * <br/> 13 = SIZEWE
     * <br/> 14 = UPARROW
     * <br/> 15 = WAIT
     */
    public int mouseGetCursor() {
        return autoItX.invoke("MouseGetCursor").getInt();
    }

    /**
     * Retrieves the current X position of the mouse cursor.
     * @return The current X position of the mouse cursor.
     */
    public int mouseGetPosX() {
        return autoItX.invoke("MouseGetPosX").getInt();
    }

    /**
     * Retrieves the current Y position of the mouse cursor.
     * @return The current Y position of the mouse cursor.
     */
    public int mouseGetPosY() {
        return autoItX.invoke("MouseGetPosY").getInt();
    }

    /**
     * Moves the mouse pointer.
     * @param x The screen x coordinate to move the mouse to.
     * @param y The screen y coordinate to move the mouse to.
     * @param speed The speed to move the mouse in the range 1 (fastest) to 
     *              100 (slowest). A speed of 0 will move the mouse instantly. 
     *              Default speed is 10.
     * @return true if success, false otherwise (is this correct)?
     */
    public boolean mouseMove(int x, int y, int speed) {
        Variant vX = new Variant(x);
        Variant vY = new Variant(y);
        Variant vSpeed = new Variant(speed);
        Variant[] params = new Variant[]{vX, vY, vSpeed};
        return oneToTrue(autoItX.invoke("MouseMove", params).getInt());
    }

    /**
     * Moves the mouse pointer.
     * @param x The screen x coordinate to move the mouse to.
     * @param y The screen y coordinate to move the mouse to.
     * @return true if success, false otherwise (is this correct)?
     */
    public boolean mouseMove(int x, int y) {
        return mouseMove(x, y, 10);
    }

    /**
     * Perform a mouse up event at the current mouse position.
     * @param button The button to click: "left", "right", "middle", "main", 
     *               "menu", "primary", "secondary".
     */
    public void mouseUp(String button) {
        autoItX.invoke("MouseUp", button);
    }

    /**
     * Moves the mouse wheel up or down. NT/2000/XP ONLY.
     * @param direction "up" or "down"
     * @param clicks The number of times to move the wheel. Default is 1.
     */
    public void mouseWheel(String direction, int clicks) {
        Variant vDirection = new Variant(direction);
        Variant vClicks = new Variant(clicks);
        Variant[] params = new Variant[]{vDirection, vClicks};
        autoItX.invoke("MouseWheel", params);
    }

    /**
     * Moves the mouse wheel up or down. NT/2000/XP ONLY.
     * @param direction "up" or "down"
     */
    public void mouseWheel(String direction) {
        mouseWheel(direction, 1);
    }

    /**
     * Terminates a named process.
     * @param process The title or PID of the process to terminate.
     */
    public void processClose(String process) {
        autoItX.invoke("ProcessClose", process);
    }

    /**
     * Checks to see if a specified process exists.
     * @param process The name or PID of the process to check.
     * @return The PID of the process. 0 if process does not exist.
     */
    public int processExists(String process) {
        return autoItX.invoke("ProcessExists", process).getInt();
    }

    /**
     * Changes the priority of a process
     * @param process The name or PID of the process to check.
     * @param priority A flag which determines what priority to set
     * <br/>    0 - Idle/Low
     * <br/>    1 - Below Normal (Not supported on Windows 95/98/ME)
     * <br/>    2 - Normal
     * <br/>    3 - Above Normal (Not supported on Windows 95/98/ME)
     * <br/>    4 - High
     * <br/>    5 - Realtime (Use with caution, may make the system unstable)
     * @return True if success, false otherwise.
     */
    public boolean processSetPriority(String process, int priority) {
        Variant vProcess = new Variant(process);
        Variant vPriority = new Variant(priority);
        Variant[] params = new Variant[]{vProcess, vPriority};
        Variant result = autoItX.invoke("ProcessSetPriority", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Pauses script execution until a given process exists.
     * @param process The name of the process to check.
     * @param timeout Specifies how long to wait (default is to wait indefinitely).
     * @return True if success, false otherwise.
     */
    public boolean processWait(String process, int timeout) {
        Variant vProcess = new Variant(process);
        Variant vTimeout = new Variant(timeout);
        Variant[] params = new Variant[]{vProcess, vTimeout};
        Variant result = autoItX.invoke("ProcessWait", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Pauses script execution until a given process exists.
     * @param process The name of the process to check.
     * @return True if success, false otherwise.
     */
    public boolean processWait(String process) {
        Variant result = autoItX.invoke("ProcessWait", process);
        return oneToTrue(result.getInt());
    }

    /**
     * Pauses script execution until a given process exists.
     * @param process The name of the process to check.
     * @param timeout Specifies how long to wait (default is to wait indefinitely).
     * @return True if success, false otherwise.
     */
    public boolean processWaitClose(String process, int timeout) {
        Variant vProcess = new Variant(process);
        Variant vTimeout = new Variant(timeout);
        Variant[] params = new Variant[]{vProcess, vTimeout};
        Variant result = autoItX.invoke("ProcessWaitClose", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Pauses script execution until a given process exists.
     * @param process The name of the process to check.
     * @return True if success, false otherwise.
     */
    public boolean processWaitClose(String process) {
        Variant result = autoItX.invoke("ProcessWaitClose", process);
        return oneToTrue(result.getInt());
    }

    /**
     * Runs an external program.
     * @param filename The name of the executable (EXE, BAT, COM, or PIF) to run.
     * @param workingDirectory The working directory.
     * @param flag The "show" flag of the executed program:
     * <br/>SW_HIDE = Hidden window
     * <br/>SW_MINIMIZE = Minimized window
     * <br/>SW_MAXIMIZE = Maximized window
     * @return The PID of the process that was launched. After running the 
     * requested program the script continues. To pause execution of the script 
     * until the spawned program has finished use the RunWait function instead.
     * The error property is set to 1 as an indication of failure.
     */
    public int run(String filename, String workingDirectory, int flag) {
        Variant vFilename = new Variant(filename);
        Variant vWorkingDirectory = new Variant(workingDirectory);
        Variant vFlag = new Variant(flag);
        Variant[] params = new Variant[]{vFilename, vWorkingDirectory, vFlag};
        return autoItX.invoke("Run", params).getInt();
    }

    /**
     * Runs an external program.
     * @param filename The name of the executable (EXE, BAT, COM, or PIF) to run.
     * @param workingDirectory The working directory.
     * @return The PID of the process that was launched. After running the 
     * requested program the script continues. To pause execution of the script 
     * until the spawned program has finished use the RunWait function instead.
     * The error property is set to 1 as an indication of failure.
     */
    public int run(String filename, String workingDirectory) {
        Variant vFilename = new Variant(filename);
        Variant vWorkingDirectory = new Variant(workingDirectory);
        Variant[] params = new Variant[]{vFilename, vWorkingDirectory};
        return autoItX.invoke("Run", params).getInt();
    }

    /**
     * Runs an external program.
     * @param filename The name of the executable (EXE, BAT, COM, or PIF) to run.
     * @return The PID of the process that was launched. After running the requested 
     * program the script continues. To pause execution of the script until the 
     * spawned program has finished use the RunWait function instead.
     * The error property is set to 1 as an indication of failure.
     */
    public int run(String filename) {
        return autoItX.invoke("Run", filename).getInt();
    }

    /**
     * Initialize a set of user credentials to use during Run and RunWait 
     * operations. 2000/XP or later ONLY.
     * @param username The user name to use.
     * @param domain The domain name to use.
     * @param password The password to use.
     * @param options 0 = do not load the user profile, 1 = (default) load the 
     *                user profile, 2 = use for net credentials only
     * @return Returns 0 if the operating system does not support this function.
     * Otherwise returns 1--regardless of success. (If the login information was 
     * invalid, subsequent Run/RunWait commands will fail....)
     */
    public int runAsSet(String username, String domain, String password, int options) {
        Variant vUsername = new Variant(username);
        Variant vDomain = new Variant(domain);
        Variant vPassword = new Variant(password);
        Variant vOptions = new Variant(options);
        Variant[] params = new Variant[]{vUsername, vDomain, vPassword, vOptions};
        return autoItX.invoke("RunAsSet", params).getInt();
    }

    /**
     * Initialize a set of user credentials to use during Run and RunWait 
     * operations. 2000/XP or later ONLY.
     * @param username The user name to use.
     * @param domain The domain name to use.
     * @param password The password to use.
     * @return Returns 0 if the operating system does not support this function.
     * Otherwise returns 1--regardless of success. (If the login information was 
     * invalid, subsequent Run/RunWait commands will fail....)
     */
    public int runAsSet(String username, String domain, String password) {
        return runAsSet(username, domain, password, 1);
    }

    /**
     * Runs an external program and pauses script execution until the program finishes.
     * @param filename The name of the executable (EXE, BAT, COM, PIF) to run.
     * @param workingDirectory The working directory.
     * @param flag The "show" flag of the executed program:
     * <br/>            SW_HIDE = Hidden window
     * <br/>            SW_MINIMIZE = Minimized window
     * <br/>            SW_MAXIMIZE = Maximized window
     * @return Returns the exit code of the program that was run. 
     * The error property is set to 1 as an indication of failure.
     */
    public int runWait(String filename, String workingDirectory, int flag) {
        Variant vFilename = new Variant(filename);
        Variant vWorkingDirectory = new Variant(workingDirectory);
        Variant vFlag = new Variant(flag);
        Variant[] params = new Variant[]{vFilename, vWorkingDirectory, vFlag};
        return autoItX.invoke("RunWait", params).getInt();
    }

    /**
     * Runs an external program and pauses script execution until the program finishes.
     * @param filename The name of the executable (EXE, BAT, COM, PIF) to run.
     * @param workingDirectory The working directory.
     * @return Returns the exit code of the program that was run. 
     * The error property is set to 1 as an indication of failure.
     */
    public int runWait(String filename, String workingDirectory) {
        Variant vFilename = new Variant(filename);
        Variant vWorkingDirectory = new Variant(workingDirectory);
        Variant[] params = new Variant[]{vFilename, vWorkingDirectory};
        return autoItX.invoke("RunWait", params).getInt();
    }

    /**
     * Runs an external program and pauses script execution until the program finishes.
     * @param filename The name of the executable (EXE, BAT, COM, PIF) to run.
     * @return Returns the exit code of the program that was run. 
     * The error property is set to 1 as an indication of failure.
     */
    public int runWait(String filename) {
        return autoItX.invoke("RunWait", filename).getInt();
    }

    /**
     * Shuts down the system.
     * @param code The shutdown code is a combination of the following values:
     * <br/>0 = Logoff
     * <br/>1 = Shutdown
     * <br/>2 = Reboot
     * <br/>4 = Force
     * <br/>8 = Power down
     * Add the required values together. To shutdown and power down, 
     * for example, the code would be 9 (shutdown + power down = 1 + 8 = 9).
     * @return True if success, false otherwise.
     */
    public boolean shutdown(int code) {
        return oneToTrue(autoItX.invoke("Shutdown", new Variant(code)).getInt());
    }

    /**
     * Deletes a key from the registry.
     * @param keyname The registry key to write to.
     * @return Success: Returns 1. Special: Returns 0 if the key does not exist. 
     * Failure: Returns 2 if error deleting key.
     */
    public int regDeleteKey(String keyname) {
        return autoItX.invoke("RegDeleteKey", keyname).getInt();
    }

    /**
     * Deletes a value from the registry.
     * @param keyname The registry key to write to.
     * @return Success: Returns 1. Special: Returns 0 if the key does not exist. 
     * Failure: Returns 2 if error deleting key.
     */
    public int regDeleteVal(String keyname) {
        return autoItX.invoke("RegDeleteVal", keyname).getInt();
    }

    /**
     * Reads the name of a subkey according to it's instance.
     * @param keyname The registry key to read.
     * @param instance The 1-based key instance to retrieve.
     * @return Success: Returns the requested subkey name. Otherwise Returns "" 
     * and sets the .error() flag,  1 if unable to open requested key, 
     * -1 if unable to retrieve requested subkey (key instance out of range)
     */
    public String regEnumKey(String keyname, int instance) {
        Variant vKeyname = new Variant(keyname);
        Variant vInstance = new Variant(instance);
        Variant[] params = new Variant[]{vKeyname, vInstance};
        return autoItX.invoke("RegEnumKey", params).getString();
    }

    /**
     * Reads the name of a value according to it's instance.
     * @param keyname The registry key to read.
     * @param instance The 1-based key instance to retrieve.
     * @return Success: Returns the requested subkey name. Otherwise Returns "" 
     * and sets the .error() flag,  1 if unable to open requested key, -1 if 
     * unable to retrieve requested subkey (key instance out of range)
     */
    public String regEnumVal(String keyname, int instance) {
        Variant vKeyname = new Variant(keyname);
        Variant vInstance = new Variant(instance);
        Variant[] params = new Variant[]{vKeyname, vInstance};
        return autoItX.invoke("RegEnumVal", params).getString();
    }

    /**
     * Reads the name of a value according to it's instance.
     * @param keyname The registry key to read.
     * @param valueName The registry value to read.
     * @return Success:Returns the requested registry value value.
     * Otherwise Returns numeric 1 and sets the oAutoIt.error flag:
     * <br/>1 if unable to open requested key
     * <br/>-1 if unable to open requested value
     * <br/>-2 if value type not supported
     */
    public String regRead(String keyname, String valueName) {
        Variant vKeyname = new Variant(keyname);
        Variant vValueName = new Variant(valueName);
        Variant[] params = new Variant[]{vKeyname, vValueName};
        Variant result = autoItX.invoke("RegRead", params);
        if (result.getvt() == Variant.VariantInt) {
            return String.valueOf(result.getInt());
        }
        
        if (result.getvt() == Variant.VariantString) {
            return result.getString();
        }
        
        return "";
    }

    /**
     * Creates a key or value in the registry.
     * @param keyname The registry key to write to. If no other parameters are 
     *                specified this key will simply be created.
     * @param valueName The valuename to write to.
     * @param type Type of key to write: "REG_SZ", "REG_MULTI_SZ", "REG_EXPAND_SZ",
     *             "REG_DWORD", or "REG_BINARY".
     * @param value The value to write.
     * @return True if success, false otherwise.
     */
    public boolean regWrite(String keyname, String valueName, String type, String value) {
        Variant vKeyname = new Variant(keyname);
        Variant vValueName = new Variant(valueName);
        Variant vType = new Variant(type);
        Variant vValue = new Variant(value);
        Variant[] params = new Variant[]{vKeyname, vValueName, vType, vValue};
        Variant result = autoItX.invoke("RegWrite", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Pause script execution.
     * @param delay Amount of time to pause (in milliseconds).
     */
    public void sleep(int delay) {
        autoItX.invoke("sleep", delay);
    }

    /**
     * Sends a mouse click command to a given control.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param controlID The control to interact with.
     * @param button The button to click, "left", "right" or "middle". 
     *               Default is the left button.
     * @param clicks The number of times to click the mouse. Default is center.
     * @param x The x position to click within the control. Default is center.
     * @param y The y position to click within the control. Default is center.
     * @return True if success, false otherwise.
     */
    public boolean controlClick(String title, String text, String controlID, 
            String button, int clicks, int x, int y) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControlID = new Variant(controlID);
        Variant vButton = new Variant(button);
        Variant vClicks = new Variant(clicks);
        Variant vX = new Variant(x);
        Variant vY = new Variant(y);
        Variant[] params = new Variant[]{vTitle, vText, vControlID, vButton, vClicks, vX, vY};
        Variant result = autoItX.invoke("ControlClick", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Sends a mouse click command to a given control.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param controlID The control to interact with.
     * @param button The button to click, "left", "right" or "middle". 
     *               Default is the left button.
     * @param clicks The number of times to click the mouse. Default is center.
     * @return True if success, false otherwise.
     */
    public boolean controlClick(String title, String text, String controlID, 
            String button, int clicks) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControlID = new Variant(controlID);
        Variant vButton = new Variant(button);
        Variant vClicks = new Variant(clicks);
        Variant[] params = new Variant[]{vTitle, vText, vControlID, vButton, vClicks};
        Variant result = autoItX.invoke("ControlClick", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Sends a mouse click command to a given control.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param controlID The control to interact with.
     * @param button The button to click, "left", "right" or "middle". 
     *               Default is the left button.
     * @return True if success, false otherwise.
     */
    public boolean controlClick(String title, String text, String controlID, String button) {
        return controlClick(title, text, controlID, button, 1);
    }

    /**
     * Sends a mouse click command to a given control.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param controlID The control to interact with.
     * @return True if success, false otherwise.
     */
    public boolean controlClick(String title, String text, String controlID) {
        return controlClick(title, text, controlID, "left", 1);
    }

    private String controlCommandString(String title, String text, String control,
            String command, String option) {
        Variant result = controlCommandVariant(title, text, control, command, option);
        return result.getString();
    }

    private void controlCommandVoid(String title, String text, String control, 
            String command, String option) {
        controlCommandVariant(title, text, control, command, option);

    }

    private Variant controlCommandVariant(String title, String text, String control, 
            String command, String option) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControl = new Variant(control);
        Variant vCommand = new Variant(command);
        Variant vOption = new Variant(option);
        Variant[] params = new Variant[]{vTitle, vText, vControl, vCommand, vOption};
        return autoItX.invoke("ControlCommand", params);
    }

    /**
     * Drops a ComboBox
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     */
    public void controlCommandShowDropdown(String title, String text, String control) {
        controlCommandVoid(title, text, control, "ShowDropDown", "");
    }

    /**
     * Undrops a ComboBox
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     */
    public void controlCommandHideDropDown(String title, String text, String control) {
        controlCommandVoid(title, text, control, "HideDropDown", "");
    }

    /**
     * Checks radio or check Button
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     */
    public void controlCommandCheck(String title, String text, String control) {
        controlCommandVoid(title, text, control, "Check", "");
    }

    /**
     * Un-checks radio or check Button
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     */
    public void controlCommandUncheck(String title, String text, String control) {
        controlCommandVoid(title, text, control, "UnCheck", "");
    }

    /**
     * Adds a string to the end in a ListBox or ComboBox
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param string The string.
     */
    public void controlCommandAddString(String title, String text, String control, String string) {
        controlCommandVoid(title, text, control, "AddString", string);
    }

    /**
     * Deletes a string according to occurrence in a ListBox or ComboBox.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param occurrance The occurrance.
     */
    public void controlCommandDeleteString(String title, String text, String control, String occurrance) {
        controlCommandVoid(title, text, control, "DelString", occurrance);
    }

    /**
     * Pastes the 'string' at the Edits caret position.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param string The string.
     */
    public void controlCommandEditPaste(String title, String text, String control, String string) {
        controlCommandVoid(title, text, control, "EditPaste", string);
    }

    /**
     * Sets selection to occurrence ref in a ListBox or ComboBox.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param occurrance
     */
    public void controlCommandSetCurrentSelection(String title, String text, 
            String control, String occurrance) {
        controlCommandVoid(title, text, control, "SetCurrentSelection", occurrance);
    }

    /**
     * Sets selection according to string in a ListBox or ComboBox
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param string The string.
     */
    public void controlCommandSelectString(String title, String text, String control, String string) {
        controlCommandVoid(title, text, control, "SelectString", string);
    }

    private boolean controlCommandBoolean(String title, String text, String control, 
            String command, String option) {
        return oneToTrue(controlCommandInts(title, text, control, command, option));
    }

    /**
     * Checks whether a control is visible.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return True if is visible.
     */
    public boolean controlCommandIsVisible(String title, String text, String control) {
        return controlCommandBoolean(title, text, control, "IsVisible", "");
    }

    /**
     * Checks whether a control is checked.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return True if is checked.
     */
    public boolean controlCommandIsChecked(String title, String text, String control) {
        return controlCommandBoolean(title, text, control, "IsChecked", "");
    }

    /**
     * Checks whether a control is enabled.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return True if is enabled.
     */
    public boolean controlCommandIsEnabled(String title, String text, String control) {
        return controlCommandBoolean(title, text, control, "IsEnabled", "");
    }

    /**
     * Returns occurrence ref of the exact string in a ListBox or ComboBox.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param string The string.
     * @return Returns occurrence ref of the exact string in a ListBox or ComboBox.
     */
    public int controlCommandFindString(String title, String text, String control, String string) {
        return controlCommandInts(title, text, control, "FindString", string);
    }

    private int controlCommandInts(String title, String text, String control, 
            String command, String option) {        
        Variant result = controlCommandVariant(title, text, control, command, option);
        return result.getvt() == Variant.VariantString
                ? Integer.parseInt(result.getString())
                : 0;
    }

    /**
     * Returns the line # where the caret is in an Edit
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return Returns the line # where the caret is in an Edit
     */
    public int controlCommandGetCurrentLine(String title, String text, String control) {
        return controlCommandInts(title, text, control, "GetCurrentLine", "");
    }

    /**
     * Returns the column # where the caret is in an Edit
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return Returns the line # where the caret is in an Edit
     */
    public int controlCommandGetCurrentCol(String title, String text, String control) {
        return controlCommandInts(title, text, control, "GetCurrentCol", "");
    }

    /**
     * Returns the column # where the caret is in an Edit
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return Returns the line # where the caret is in an Edit
     */
    public int controlCommandGetLineCount(String title, String text, String control) {
        return controlCommandInts(title, text, control, "GetLineCount", "");
    }

    /**
     * Sets selection according to string in a ListBox or ComboBox
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param charLength The char size.
     * @return The current selection.
     */
    public String controlCommandGetCurrentSelection(String title, String text, 
            String control, int charLength) {
        return controlCommandString(title, text, control, "GetCurrentSelection", "");
    }

    /**
     * Returns selected text of an Edit.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param charLength The length to get.
     * @return The selected text.
     */
    public String controlCommandGetSelected(String title, String text, String control, int charLength) {
        return controlCommandString(title, text, control, "GetSelected", "");
    }

    /**
     * Moves to the next tab to the right of a SysTabControl32
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     */
    public void controlCommandTabLeft(String title, String text, String control) {
        controlCommandVoid(title, text, control, "TabLeft", "");
    }

    /**
     * Moves to the next tab to the right of a SysTabControl32
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     */
    public void controlCommandTabRight(String title, String text, String control) {
        controlCommandVoid(title, text, control, "TabRight", "");
    }

    /**
     * Returns the current Tab shown of a SysTabControl32.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return Returns the current Tab shown of a SysTabControl32
     */
    public String controlCommandCurrentTab(String title, String text, String control) {
        return controlCommandString(title, text, control, "CurrentTab", "");
    }

    /**
     * Disables or "grays-out" a control.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return True if success, false otherwise.
     */
    public boolean controlDisable(String title, String text, String control) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControl = new Variant(control);
        Variant[] params = new Variant[]{vTitle, vText, vControl};
        Variant result = autoItX.invoke("ControlDisable", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Enables a "grayed-out" control.
     * <br/><b>Use with caution.</b>
     * <br/>
     * When using a control name in the Control functions, you need to add a 
     * number to the end of the name to indicate which control. For example, 
     * if there two controls listed called "MDIClient", you would refer to 
     * these as "MDIClient1" and "MDIClient2".
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return True if success, false otherwise.
     */
    public boolean controlEnable(String title, String text, String control) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControl = new Variant(control);
        Variant[] params = new Variant[]{vTitle, vText, vControl};
        Variant result = autoItX.invoke("ControlEnable", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Sets input focus to a given control on a window.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return True if success, false otherwise.
     */
    public boolean controlFocus(String title, String text, String control) {
        return controlBool(title, text, control, "ControlFocus");
    }

    /**
     * Returns the ControlRef# of the control that has keyboard focus within 
     * a specified window.
     * @param title Title of window to check.
     * @param text Text from window to check.
     * @return ControlRef# of the control that has keyboard focus within a 
     * specified window. Otherwise returns a blank string and sets .error() 
     * to 1 if window is not found.
     */
    public String controlGetFocus(String title, String text) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant[] params = new Variant[]{vTitle, vText};
        return autoItX.invoke("ControlGetFocus", params).getString();
    }

    /**
     * Returns the ControlRef# of the control that has keyboard focus within a 
     * specified window.
     * @param title Title of window to check.
     * @return ControlRef# of the control that has keyboard focus within a 
     * specified window. Otherwise returns a blank string and sets .error() 
     * to 1 if window is not found.
     */
    public String controlGetFocus(String title) {
        return autoItX.invoke("ControlGetFocus", title).getString();
    }

    /**
     * Retrieves the internal handle of a control.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @param controlID The control to interact with.
     * @return Returns a string containing the control handle value. 
     * Otherwise returns "" (blank string) and sets oAutoIt.error to 1 
     * if no window matches the criteria.
     */
    public String controlGetHandle(String title, String text, String controlID) {
        return controlString(title, text, controlID, "ControlGetHandle");
    }

    /**
     * Retrieves the position and size of a control relative to it's window.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param controlID The control to interact with.
     * @return The X coordinate of the control. 
     */
    public int controlGetPosX(String title, String text, String controlID) {
        return controlInt(title, text, controlID, "ControlGetPosX");
    }

    /**
     * Retrieves the position and size of a control relative to it's window.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param controlID The control to interact with.
     * @return The Y coordinate of the control. 
     */
    public int controlGetPosY(String title, String text, String controlID) {
        return controlInt(title, text, controlID, "ControlGetPosY");
    }

    /**
     * Retrieves the position and size of a control relative to it's window.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param controlID The control to interact with.
     * @return The width of the control. Otherwise sets .error() to 1.
     */
    public int controlGetPosWidth(String title, String text, String controlID) {
        return controlInt(title, text, controlID, "ControlGetPosWidth");
    }

    /**
     * Retrieves the position and size of a control relative to it's window.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param controlID The control to interact with.
     * @return Returns the height of the control. 
     * Otherwise returns the Y coordinate of the control.
     */
    public int controlGetPosHeight(String title, String text, String controlID) {
        return controlInt(title, text, controlID, "ControlGetPosHeight");
    }

    /**
     * Retrieves text from a control.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param controlID The control to interact with.
     * @return Text from a control. 
     * If fails, sets .error() to 1 and returns a blank string of "".
     */
    public String controlGetText(String title, String text, String controlID) {
        return controlString(title, text, controlID, "ControlGetText");
    }

    /**
     * Hides a control.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param controlID The control to interact with.
     * @return True if success, false otherwise.
     */
    public boolean controlHide(String title, String text, String controlID) {
        return controlBool(title, text, controlID, "ControlHide");
    }

    /**
     * Returns the item index of the string. Returns -1 if the string is not found.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param string String to find
     * @param subitem The SubItem
     * @return Returns the item index of the string. Returns -1 if the string is not found.
     */
    public int controlListViewFindItem(String title, String text, String control, 
            String string, String subitem) {
        return controlListViewInt(title, text, control, "FindItem", string, subitem);
    }

    private int controlListViewInt(String title, String text, String control, 
            String command, String option, String option2) {
        return controlView(title, text, control, command, option, option2, "ControlListView").getInt();
    }

    private String controlListViewString(String title, String text, String control, 
            String command, String option, String option2) {
        return controlView(title, text, control, command, option, option2, "ControlListView").getString();
    }

    private Variant controlView(String title, String text, String control, 
            String command, String option, String option2, String function) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControl = new Variant(control);
        Variant vCommand = new Variant(command);
        Variant vOption = new Variant(option);
        Variant vOption2 = new Variant(option2);
        Variant[] params = new Variant[]{vTitle, vText, vControl, vCommand, vOption, vOption2};
        return autoItX.invoke(function, params);
    }

    /**
     * Returns the number of list items.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return Returns the number of list items.
     */
    public int controlListViewGetItemCount(String title, String text, String control) {
        return controlListViewInt(title, text, control, "GetItemCount", "", "");

    }

    /**
     * Returns the number of items that are selected.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return Returns the number of items that are selected.
     */
    public int controlListViewGetSelectedCount(String title, String text, String control) {
        return controlListViewInt(title, text, control, "GetSelectedCount", "", "");
    }

    /**
     * Returns the number of subitems.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return Returns the number of subitems.
     */
    public int controlListViewGetSubItemCount(String title, String text, String control) {
        return controlListViewInt(title, text, control, "GetSubItemCount", "", "");
    }

    /**
     * Returns the text of a given item/subitem.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param item The text of an item.
     * @param subitem The text of a subitem.
     * @return Returns the text of a given item/subitem.
     */
    public String controlListViewGetText(String title, String text, String control, 
            String item, String subitem) {
        return controlListViewString(title, text, control, "GetText", item, subitem);
    }

    /**
     * Returns 1 if the item is selected, otherwise returns 0.
     * Returns the text of a given item/subitem.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param item The text of an item.
     * @return Returns 1 if the item is selected, otherwise returns 0.
     */
    public boolean controlListViewIsSelected(String title, String text, String control, String item) {
        return oneToTrue(controlListViewInt(title, text, control, "IsSelected", item, ""));
    }

    /**
     * Returns a string containing the item index of selected items. 
     * If option=0 (default) only the first selected item is returned. 
     * If option=1 then all the selected items are returned delimited by |, 
     * e.g: "0|3|4|10". If no items are selected a blank "" string is returned.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return Returns a string containing the item index of selected items. 
     * If option=0 (default) only the first selected item is returned. 
     * If option=1 then all the selected items are returned delimited by |, 
     * e.g: "0|3|4|10". If no items are selected a blank "" string is returned.
     */
    public String controlListViewGetSelected(String title, String text, String control) {
        return controlListViewString(title, text, control, "GetSelected", "", "");
    }

    /**
     * Returns a string containing the item index of selected items. 
     * If option=0 (default) only the first selected item is returned. 
     * If option=1 then all the selected items are returned delimited by |, 
     * e.g: "0|3|4|10". If no items are selected a blank "" string is returned.
     * If no items are selected a blank "" string is returned.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return Returns a string containing the item index of selected items. 
     * If option=0 (default) only the first selected item is returned. 
     * If option=1 then all the selected items are returned delimited by |, 
     * e.g: "0|3|4|10". If no items are selected a blank "" string is returned.
     */
    public String[] controlListViewGetSelectedArray(String title, String text, String control) {
        SafeArray safeArr = controlView(title, text, control, "GetSelected", "", "", "ControlListView").toSafeArray();
        return safeArr.toStringArray();
    }

    /**
     * Selects one or more items.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param from The start index.
     * @param to The to index.
     */
    public void controlListViewSelect(String title, String text, String control, 
            String from, String to) {
        controlView(title, text, control, "Select", from, to, "ControlListView");
    }

    /**
     * Selects all items.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param from The start index.
     * @param to The end index. 
     */
    public void controlListViewSelectAll(String title, String text, String control, 
            String from, String to) {
        controlView(title, text, control, "SelectAll", from, to, "ControlListView");
    }

    /**
     * Selects all items.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param from The start index.
     */
    public void controlListViewSelectAll(String title, String text, String control, String from) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControl = new Variant(control);
        Variant vCommand = new Variant("SelectAll");
        Variant vFrom = new Variant(from);
        Variant[] params = new Variant[]{vTitle, vText, vControl, vCommand, vFrom};
        autoItX.invoke("ControlListView", params);
    }

    /**
     * Clears the selection of all items.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     */
    public void controlListViewSelectClear(String title, String text, String control) {
        controlView(title, text, control, "SelectClear", "", "", "ControlListView");
    }

    /**
     * Inverts the current selection.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     */
    public void controlListViewSelectInvert(String title, String text, String control) {
        controlView(title, text, control, "SelectInvert", "", "", "ControlListView");
    }

    /**
     * Changes the current view. Valid views are "list", "details", "smallicons", 
     * "largeicons".
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param view Valid views are "list", "details", "smallicons", "largeicons".
     */
    public void controlListViewSelectViewChange(String title, String text, 
            String control, String view) {
        controlView(title, text, control, "ViewChnage", view, "", "ControlListView");
    }

    protected Variant controlVariant(String title, String text, String control, String function) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControl = new Variant(control);
        Variant[] params = new Variant[]{vTitle, vText, vControl};
        return autoItX.invoke(function, params);
    }

    protected boolean controlBool(String title, String text, String control, String function) {
        Variant result = controlVariant(title, text, control, function);
        return oneToTrue(result.getInt());
    }

    protected int controlInt(String title, String text, String control, String function) {
        Variant result = controlVariant(title, text, control, function);
        return result.getInt();
    }

    protected String controlString(String title, String text, String control, String function) {
        Variant result = controlVariant(title, text, control, function);
        if (result.getvt() == Variant.VariantString) {
            return result.getString();
        }
        
        if (result.getvt() == Variant.VariantInt) {
            return String.valueOf(result.getInt());
        }
        
        return "";
    }

    /**
     * Moves a control within a window.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param x X coordinate to move to.
     * @param y Y coordinate to move to.
     * @param width New width of the window.
     * @param height New height of the window.
     * @return True if success, false otherwise
     */
    public boolean controlMove(String title, String text, String control, 
            int x, int y, int width, int height) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControl = new Variant(control);
        Variant vX = new Variant(x);
        Variant vY = new Variant(y);
        Variant vWidth = new Variant(width);
        Variant vHeight = new Variant(height);
        Variant[] params = new Variant[]{vTitle, vText, vControl, vX, vY, vWidth, vHeight};
        Variant result = autoItX.invoke("ControlMove", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Moves a control within a window.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param x X coordinate to move to.
     * @param y Y coordinate to move to.
     * @return True if success, false otherwise
     */
    public boolean controlMove(String title, String text, String control, int x, int y) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControl = new Variant(control);
        Variant vX = new Variant(x);
        Variant vY = new Variant(y);
        Variant[] params = new Variant[]{vTitle, vText, vControl, vX, vY};
        Variant result = autoItX.invoke("ControlMove", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Sends a string of characters to a control.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param string String of characters to send to the control.
     * @param sendRawKeys If true, text contains special characters like + 
     *                    to indicate SHIFT and {LEFT} to indicate left arrow. 
     *                    If false, text is sent raw.
     * @return  True if success, false otherwise
     */
    public boolean controlSend(String title, String text, String control, 
            String string, boolean sendRawKeys) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControl = new Variant(control);
        Variant vString = new Variant(string);
        int flag = (sendRawKeys) ? 1 : 0;
        Variant vFlag = new Variant(flag);
        Variant[] params = new Variant[]{vTitle, vText, vControl, vString, vFlag};
        Variant result = autoItX.invoke("ControlSend", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Sends a string of characters to a control.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param string String of characters to send to the control.
     * @return True if success, false otherwise
     */
    public boolean controlSend(String title, String text, String control, String string) {
        return controlSend(title, text, control, string, false);
    }

    /**
     * Sets text of a control.
     * Sends a string of characters to a control.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param string The new text to be set into the control.
     * @return True if success, false otherwise
     */
    public boolean ControlSetText(String title, String text, String control, String string) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControl = new Variant(control);
        Variant vString = new Variant(string);
        Variant[] params = new Variant[]{vTitle, vText, vControl, vString};
        Variant result = autoItX.invoke("ControlSetText", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Shows a control that was hidden.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return True if success, false otherwise.
     */
    public boolean controlShow(String title, String text, String control) {
        return controlBool(title, text, control, "ControlShow");
    }

    private String controlTreeViewString(String title, String text, String control, 
            String command, String option, String option2) {
        return controlView(title, text, control, command, option, option2, "ControlTreeView").getString();
    }

    private int controlTreeViewInt(String title, String text, String control, 
            String command, String option, String option2) {
        return controlView(title, text, control, command, option, option2, "ControlTreeView").getInt();
    }

    public boolean controlTreeViewBoolean(String title, String text, String control, 
            String command, String option, String option2) {
        Variant result = controlView(title, text, control, command, option, option2, "ControlTreeView");
        return oneToTrue(result.getInt());
    }

    /**
     * Checks an item (if the item supports it).
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param item The item to check
     */
    public void controlTreeViewCheck(String title, String text, String control, String item) {
        controlView(title, text, control, "Check", item, "", "ControlTreeView");
    }

    /**
     * Collapses an item to hide its children.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param item The item to check
     */
    public void controlTreeViewCollapse(String title, String text, String control, String item) {
        controlView(title, text, control, "Collapse", item, "", "ControlTreeView");
    }

    /**
     * Checks if an item exists
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param item The item to check
     * @return True if the tree view exists.
     */
    public Boolean controlTreeViewExists(String title, String text, String control, String item) {
        return controlTreeViewBoolean(title, text, control, "Exists", item, "");
    }

    /**
     * Expands an item to show its children.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param item The item to expand
     */
    public void controlTreeViewExpand(String title, String text, String control, String item) {
        controlView(title, text, control, "Expand", item, "", "ControlTreeView");
    }

    /**
     * Returns the number of children for a selected item.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param item The item to check
     * @return The number of children for a selected item.
     */
    public int controlTreeViewGetItemCount(String title, String text, String control, String item) {
        return controlTreeViewInt(title, text, control, "GetItemCount", item, "");
    }

    /**
     * Returns the item reference of the current selection using the index reference of the item.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return The item reference of the current selection using the index reference of the item.
     */
    public int controlTreeViewGetSelectedItemIndex(String title, String text, String control) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControl = new Variant(control);
        Variant vCommand = new Variant("GetSelected");
        Variant vIndex = new Variant(1);
        Variant[] params = new Variant[]{vTitle, vText, vControl, vCommand, vIndex};
        return autoItX.invoke("ControlTreeView", params).getInt();
    }

    /**
     * Returns the item reference of the current selection using the text reference of the item.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return The item reference of the current selection using the text reference of the item.
     */
    public String controlTreeViewGetSelectedItemText(String title, String text, String control) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vControl = new Variant(control);
        Variant vCommand = new Variant("GetSelected");
        Variant vIndex = new Variant(0);
        Variant[] params = new Variant[]{vTitle, vText, vControl, vCommand, vIndex};
        return autoItX.invoke("ControlTreeView", params).getString();
    }

    /**
     * GetText
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param item The item to get
     * @return The the text of an item.
     */
    public String controlTreeViewGetText(String title, String text, String control, String item) {
        return controlTreeViewString(title, text, control, "GetText", item, "");
    }

    /**
     * Returns the state of an item. 1:checked, 0:unchecked, -1:not a checkbox.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @return Returns the state of an item. 1:checked, 0:unchecked, -1:not a checkbox.
     */
    public int controlTreeViewIsChecked(String title, String text, String control) {
        return controlView(title, text, control, "IsChecked", "", "", "ControlTreeView").getInt();
    }

    /**
     * Selects an item.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param item The item to select
     */
    public void controlTreeViewSelect(String title, String text, String control, String item) {
        controlView(title, text, control, "Select", item, "", "ControlTreeView");
    }

    /**
     * Uncheck an item.
     * @param title The title of the window to access.
     * @param text The text of the window to access.
     * @param control The control to interact with.
     * @param item The item to select
     */
    public void controlTreeViewUncheck(String title, String text, String control, String item) {
        controlView(title, text, control, "Uncheck", item, "", "ControlTreeView");
    }

    /**
     * Retrieves the text from a standard status bar control.
     * @param title The title of the window to check.
     * @param text The text of the window to check.
     * @param part The "part" number of the status bar to read 
     *             the default is 1. 1 is the first possible part and usually 
     *             the one that contains the useful messages like "Ready" "Loading...", etc.
     * @return The text from a standard status bar control.
     */
    public String statusbarGetText(String title, String text, int part) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vPart = new Variant(part);
        Variant[] params = new Variant[]{vTitle, vText, vPart};
        return autoItX.invoke("StatusbarGetText", params).getString();
    }

    /**
     * Retrieves the text from a standard status bar control.
     * @param title The title of the window to check.
     * @param text The text of the window to check.
     * @return  The text from a standard status bar control.
     */
    public String StatusbarGetText(String title, String text) {
        return winVariant(title, text, "StatusbarGetText").getString();

    }

    private Variant winVariant(String title, String text, String function) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant[] params = new Variant[]{vTitle, vText};
        return autoItX.invoke(function, params);
    }

    private Variant winVariant(String title, String function) {
        Variant vTitle = new Variant(title);
        Variant[] params = new Variant[]{vTitle};
        return autoItX.invoke(function, params);
    }

    /**
     * Activates (gives focus to) a window.
     * @param title The title of the window to activate.
     * @param text The text of the window to activate.
     */
    public void winActivate(String title, String text) {
        winVariant(title, text, "WinActivate");
    }

    /**
     * Activates (gives focus to) a window.
     * @param title The title of the window to activate.
     */
    public void winActivate(String title) {
        winVariant(title, "WinActivate");
    }

    /**
     * Checks to see if a specified window exists and is currently active.
     * @param title The title of the window to activate.
     * @param text The text of the window to activate.
     */
    public void winActive(String title, String text) {
        winVariant(title, text, "WinActive");
    }

    /**
     * Checks to see if a specified window exists and is currently active.
     * @param title The title of the window to activate.
     */
    public void winActive(String title) {
        winVariant(title, "WinActive");
    }

    /**
     * Closes a window.
     * @param title The title of the window to activate.
     * @param text The text of the window to activate.
     */
    public void winClose(String title, String text) {
        winVariant(title, text, "WinClose");
    }

    /**
     * Closes a window.
     * @param title The title of the window to activate.
     */
    public void winClose(String title) {
        winVariant(title, "WinClose");
    }

    /**
     * Checks to see if a specified window exists.
     * @param title The title of the window to activate.
     * @param text The text of the window to activate.
     * @return True if window exists, false otherwise.
     */
    public boolean winExists(String title, String text) {
        Variant result = winVariant(title, text, "WinExists");
        return oneToTrue(result.getInt());
    }

    /**
     * Checks to see if a specified window exists.
     * @param title The title of the window to activate.
     * @return True if window exists, false otherwise.
     */
    public boolean winExists(String title) {
        Variant result = winVariant(title, "WinExists");
        return oneToTrue(result.getInt());
    }

    /**
     * Returns the coordinates of the caret in the foreground window
     * @return The coordinates of the caret in the foreground window
     */
    public int winGetCaretPosX() {
        return autoItX.invoke("WinGetCaretPosX").getInt();
    }

    /**
     * Returns the coordinates of the caret in the foreground window
     * @return The coordinates of the caret in the foreground window
     */
    public int winGetCaretPosY() {
        return autoItX.invoke("WinGetCaretPosY").getInt();
    }

    /**
     * Retrieves the classes from a window.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @return A string containing the window classes read, otherwise returns 
     * empty string and sets .error() to 1.
     */
    public String winGetClassList(String title, String text) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant[] params = new Variant[]{vTitle, vText};
        Variant result = autoItX.invoke("WinGetClassList", params);
        return safeString(result);
    }

    /**
     * Retrieves the size of a given window's client area.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @return Returns the width of the window's client area, else returns 1 
     * and sets .error() =1;
     */
    public int winGetClientSizeWidth(String title, String text) {
        Variant result = winVariant(title, text, "WinGetClientSizeWidth");
        return result.getInt();
    }

    /**
     * Retrieves the size of a given window's client area.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @return Returns the height of the window's client area, else returns 1 
     * and sets .error() =1;
     */
    public int winGetClientSizeHeight(String title, String text) {
        Variant result = winVariant(title, text, "WinGetClientSizeHeight");
        return result.getInt();
    }

    /**
     * Retrieves the size of a given window's client area.
     * @param title The title of the window to read.
     * @return Returns the width of the window's client area, else returns 1 
     * and sets .error() =1;
     */
    public int winGetClientSizeWidth(String title) {
        Variant result = winVariant(title, "WinGetClientSizeWidth");
        return result.getInt();
    }

    /**
     * Retrieves the size of a given window's client area.
     * @param title The title of the window to read.
     * @return Returns the height of the window's client area, else returns 1 
     * and sets .error() =1;
     */
    public int winGetClientSizeHeight(String title) {
        Variant result = winVariant(title, "WinGetClientSizeHeight");
        return result.getInt();
    }

    private String safeString(Variant v) {
        return v.getvt() == Variant.VariantString ? v.getString() : "";
    }

    /**
     * Retrieves the internal handle of a window.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @return A string containing the window handle value. Otherwise returns "" 
     * and sets .error() to 1.
     */
    public String winGetHandle(String title, String text) {
        Variant result = winVariant(title, text, "WinGetHandle");
        return result.getString();
    }

    /**
     * Retrieves the internal handle of a window.
     * @param title The title of the window to read.
     * @return A string containing the window handle value. Otherwise returns "" 
     * and sets .error() to 1.
     */
    public String winGetHandle(String title) {
        Variant result = winVariant(title, "WinGetHandle");
        return result.getString();
    }

    /**
     * Retrieves the position and size of a given window.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @return Returns the X coordinate of the window. Otherwise returns 1 and 
     * sets .error() = 1
     */
    public int winGetPosX(String title, String text) {
        return winVariant(title, text, "WinGetPosX").getInt();
    }

    /**
     * Retrieves the position and size of a given window.
     * @param title The title of the window to read.
     * @return Returns the X coordinate of the window. Otherwise returns 1 and 
     * sets .error() = 1
     */
    public int winGetPosX(String title) {
        return winVariant(title, "WinGetPosX").getInt();
    }

    /**
     * Retrieves the position and size of a given window.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @return Returns the Y coordinate of the window. Otherwise returns 1 and 
     * sets .error() = 1
     */
    public int winGetPosY(String title, String text) {
        Variant result = winVariant(title, text, "WinGetPosY");
        return result.getInt();
    }

    /**
     * Retrieves the position and size of a given window.
     * @param title The title of the window to read.
     * @return Returns the Y coordinate of the window. Otherwise returns 1 and 
     * sets .error() = 1
     */
    public int winGetPosY(String title) {
        Variant result = winVariant(title, "WinGetPosY");
        return result.getInt();
    }

    /**
     * Retrieves the position and size of a given window.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @return Returns the width of the window. Otherwise returns 1 and sets 
     * error() = 1
     */
    public int winGetPosWidth(String title, String text) {
        Variant result = winVariant(title, text, "WinGetPosWidth");
        return result.getInt();
    }

    /**
     * Retrieves the position and size of a given window.
     * @param title The title of the window to read.
     * @return Returns the width of the window. Otherwise returns 1 and sets 
     * error() = 1
     */
    public int winGetPosWidth(String title) {
        Variant result = winVariant(title, "WinGetPosWidth");
        return result.getInt();
    }

    /**
     * Retrieves the position and size of a given window.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @return Returns the height of the window. Otherwise returns 1 and sets 
     * error() = 1
     */
    public int winGetPosHeight(String title, String text) {
        Variant result = winVariant(title, text, "WinGetPosHeight");
        return result.getInt();
    }

    /**
     * Retrieves the position and size of a given window.
     * @param title The title of the window to read.
     * @return Returns the height of the window. Otherwise returns 1 and sets 
     * error() = 1
     */
    public int winGetPosHeight(String title) {
        Variant result = winVariant(title, "WinGetPosHeight");
        return result.getInt();
    }

    /**
     * Retrieves the Process ID (PID) associated with a window.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @return The PID, otherwise returns "".
     */
    public String winGetProcess(String title, String text) {
        Variant v = winVariant(title, text, "WinGetProcess");
        return v.getString();
    }

    /**
     * Retrieves the Process ID (PID) associated with a window.
     * @param title The title of the window to read.
     * @return The PID, otherwise returns "".
     */
    public String winGetProcess(String title) {
        Variant v = winVariant(title, "WinGetProcess");
        return v.getString();
    }

    /**
     * Retrieves the state of a given window.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @return  Returns a value indicating the state of the window. 
     * Multiple values are added together so use BitAND() to examine the part 
     * you are interested in:
     * <br/>1 =  Window exists
     * <br/>2 = Window is visible
     * <br/>4 = Windows is enabled
     * <br/>8 = Window is active
     * <br/>16 = Window is minimized
     * <br/>otherwise returns 0 and sets oAutoIt.error to 1 if the window is not found.
     */
    public int winGetState(String title, String text) {
        Variant result = winVariant(title, text, "WinGetState");
        return result.getInt();
    }

    /**
     * Retrieves the state of a given window.
     * @param title The title of the window to read.
     * @return  Returns a value indicating the state of the window. 
     * Multiple values are added together so use BitAND() to examine the 
     * part you are interested in:
     * <br/> 1 =  Window exists
     * <br/> 2 = Window is visible
     * <br/> 4 = Windows is enabled
     * <br/> 8 = Window is active
     * <br/> 16 = Window is minimized
     * <br/> otherwise returns 0 and sets oAutoIt.error to 1 if the window is not found.
     */
    public int winGetState(String title) {
        Variant result = winVariant(title, "WinGetState");
        return result.getInt();
    }

    /**
     * Retrieves up to 64KB of text from a window.
     * <p>WinGetText works on minimized windows, but only works on hidden windows 
     * if you've set AutoItSetOption("WinDetectHiddenText", 1)
     * If multiple windows match the criteria for WinGetText, the information 
     * for the most recently active match is returned.
     * Use WinGetText("") to get the active window's text.</p>
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @return Returns a string containing up to 64k of the window text read.
     */
    public String winGetText(String title, String text) {
        Variant result = winVariant(title, text, "WinGetText");
        return result.getString();
    }

    /**
     * Retrieves up to 64KB of text from a window.
     * <p>WinGetText works on minimized windows, but only works on hidden windows 
     * if you've set AutoItSetOption("WinDetectHiddenText", 1)
     * If multiple windows match the criteria for WinGetText, the information 
     * for the most recently active match is returned.
     * Use WinGetText("") to get the active window's text.</p>
     * @param title The title of the window to read.
     * @return Returns a string containing up to 64k of the window text read.
     */
    public String winGetText(String title) {
        Variant result = winVariant(title, "WinGetText");
        return result.getString();
    }

    /**
     * Retrieves the full title from a window.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @return  A string containing the complete window title. Otherwise "".
     */
    public String winGetTitle(String title, String text) {
        Variant result = winVariant(title, text, "WinGetTitle");
        if (result.getvt() == Variant.VariantString) {
            return result.getString();
        }
        return "";
    }

    /**
     * Retrieves the full title from a window.
     * @param title The title of the window to read.
     * @return  A string containing the complete window title. Otherwise "".
     */
    public String winGetTitle(String title) {
        Variant result = winVariant(title, "WinGetTitle");
        if (result.getvt() == Variant.VariantString) {
            return result.getString();
        }
        return "";
    }

    /**
     * Forces a window to close.
     * @param title The title of the window.
     * @param text The text of the window.
     */
    public void winKill(String title, String text) {
        winVariant(title, text, "WinKill");
    }

    /**
     * Forces a window to close.
     * @param title The title of the window.
     */
    public void winKill(String title) {
        winVariant(title, "WinKill");
    }

    /**
     * Retrieves a list of windows.
     * @param title The title of the window.
     * @param text The text of the window.
     * @return Returns a 2 dimensional array containing the window titles 
     * and corresponding handles.
     */
    public String[][] winList(String title, String text) {
        Variant result = winVariant(title, text, "WinList");
        SafeArray arr = result.toSafeArray();
        int entries = arr.getInt(0, 0);
        String[][] resultArr = new String[2][entries + 1];
        for (int i = 0; i <= entries; i++) {
            resultArr[0][i] = arr.getString(0, i);
            resultArr[1][i] = arr.getString(1, i);
        }
        return resultArr;
    }

    /**
     * Retrieves a list of windows.
     * @param title The title of the window.
     * @return Returns a 2 dimensional array containing the window titles 
     * and corresponding handles.
     */
    public String[][] winList(String title) {
        Variant result = winVariant(title, "WinList");
        SafeArray arr = result.toSafeArray();
        int entries = arr.getInt(0, 0);
        String[][] resultArr = new String[2][entries + 1];
        for (int i = 0; i <= entries; i++) {
            resultArr[0][i] = arr.getString(0, i);
            resultArr[1][i] = arr.getString(1, i);
        }
        
        return resultArr;
    }

    /**
     * Invokes a menu item of a window.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @param item Text of Menu Item
     * @return True if success, false otherwise.
     */
    public boolean winMenuSelectItem(String title, String text, String item) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vItem = new Variant(item);
        Variant[] params = new Variant[]{vTitle, vText, vItem};
        Variant result = autoItX.invoke("WinMenuSelectItem", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Invokes a menu item of a window.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @param item Text of Menu Item
     * @param item2 Text of Menu Item
     * @return True if success, false otherwise.
     */
    public boolean winMenuSelectItem(String title, String text, String item, String item2) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vItem = new Variant(item);
        Variant vItem2 = new Variant(item2);
        Variant[] params = new Variant[]{vTitle, vText, vItem, vItem2};
        Variant result = autoItX.invoke("WinMenuSelectItem", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Invokes a menu item of a window.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @param item Text of Menu Item
     * @param item2 Text of Menu Item
     * @param item3 Text of Menu Item
     * @return True if success, false otherwise.
     */
    public boolean winMenuSelectItem(String title, String text, String item, 
            String item2, String item3) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vItem = new Variant(item);
        Variant vItem2 = new Variant(item2);
        Variant vItem3 = new Variant(item3);
        Variant[] params = new Variant[]{vTitle, vText, vItem, vItem2, vItem3};
        Variant result = autoItX.invoke("WinMenuSelectItem", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Invokes a menu item of a window.
     * @param title The title of the window to read.
     * @param text The text of the window to read.
     * @param item Text of Menu Item
     * @param item2 Text of Menu Item
     * @param item3 Text of Menu Item
     * @param item4 Text of Menu Item
     * @param item5 Text of Menu Item
     * @param item6 Text of Menu Item
     * @param item7 Text of Menu Item
     * @return True if success, false otherwise.
     */
    public boolean winMenuSelectItem(String title, String text, String item, 
            String item2, String item3, String item4, String item5, String item6, String item7) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vItem = new Variant(item);
        Variant vItem2 = new Variant(item2);
        Variant vItem3 = new Variant(item3);
        Variant vItem4 = new Variant(item4);
        Variant vItem5 = new Variant(item5);
        Variant vItem6 = new Variant(item6);
        Variant vItem7 = new Variant(item7);
        Variant[] params = new Variant[]{vTitle, vText, vItem, vItem2, vItem3, vItem4, vItem5, vItem6, vItem7};
        Variant result = autoItX.invoke("WinMenuSelectItem", params);
        return oneToTrue(result.getInt());
    }

    /**
     * Minimizes all windows.
     */
    public void winMinimizeAll() {
        autoItX.invoke("WinMinimizeAll");
    }

    /**
     * Undoes a previous WinMinimizeAll function.
     */
    public void winMinimizeAllUndo() {
        autoItX.invoke("WinMinimizeAllUndo");
    }

    /**
     * Moves and/or resizes a window.
     * @param title The title of the window to move/resize.
     * @param text The text of the window to move/resize.
     * @param x X coordinate to move to.
     * @param y Y coordinate to move to.
     * @param width New width of the window.
     * @param height New height of the window.
     */
    public void winMove(String title, String text, int x, int y, int width, int height) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vX = new Variant(x);
        Variant vY = new Variant(y);
        Variant vWidth = new Variant(width);
        Variant vHeight = new Variant(height);
        Variant[] params = new Variant[]{vTitle, vText, vX, vY, vWidth, vHeight};
        autoItX.invoke("WinMove", params);
    }

    /**
     * Moves and/or resizes a window.
     * @param title The title of the window to move/resize.
     * @param text The text of the window to move/resize.
     * @param x X coordinate to move to.
     * @param y Y coordinate to move to.
     */
    public void winMove(String title, String text, int x, int y) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vX = new Variant(x);
        Variant vY = new Variant(y);
        Variant[] params = new Variant[]{vTitle, vText, vX, vY};
        autoItX.invoke("WinMove", params);
    }

    /**
     * Change a window's "Always On Top" attribute.
     * @param title The title of the window to affect.
     * @param text The text of the window to affect.
     * @param isTopMost Determines whether the window should have the "TOPMOST" flag set.
     * true=set on top flag, false = remove on top flag
     */
    public void winSetOnTop(String title, String text, boolean isTopMost) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        int flag = 0;
        if (isTopMost) {
            flag = 1;
        }
        Variant vFlag = new Variant(flag);
        Variant[] params = new Variant[]{vTitle, vText, vFlag};
        autoItX.invoke("WinSetOnTop", params);
    }

    /**
     * Shows, hides, minimizes, maximizes, or restores a window.
     * @param title The title of the window to affect.
     * @param text The text of the window to affect.
     * @param flag The "show" flag of the executed program:
     * <br/>SW_HIDE = Hide window,
     * <br/>SW_SHOW = Shows a previously hidden window,
     * <br/>SW_MINIMIZE = Minimize window,
     * <br/>SW_MAXIMIZE = Maximize window,
     * <br/>SW_RESTORE = Undoes a window minimization or maximization.
     */
    public void winSetState(String title, String text, int flag) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vFlag = new Variant(flag);
        Variant[] params = new Variant[]{vTitle, vText, vFlag};
        autoItX.invoke("WinSetState", params);
    }

    /**
     * Changes the title of a window.
     * @param title The title of the window to affect.
     * @param text The text of the window to affect.
     * @param newtitle The new title to give to the window.
     */
    public void winSetTitle(String title, String text, String newtitle) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vNewtitle = new Variant(newtitle);
        Variant[] params = new Variant[]{vTitle, vText, vNewtitle};
        autoItX.invoke("WinSetTitle", params);
    }

    /**
     * Sets the transparency of a window. (Windows 2000/XP or later)
     * @param title The title of the window to affect.
     * @param text The text of the window to affect.
     * @param transparency A number in the range 0 - 255. The larger the number, 
     *                     the more transparent the window will become.
     * @return True on success, false on failure. .error() will be set to 1 
     * if the function isn't supported on an OS.
     */
    public boolean winSetTrans(String title, String text, int transparency) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vTransparency = new Variant(transparency);
        Variant[] params = new Variant[]{vTitle, vText, vTransparency};
        Variant result = autoItX.invoke("WinSetTrans", params);
        return result.getInt() != 0;
    }

    /**
     * Pauses execution of the script until the requested window exists.
     * The script polls for window match every 250 milliseconds or so.
     * @param title The title of the window to check.
     * @param text The text of the window to check.
     * @param timeout Timeout in seconds
     * @return True if success, false otherwise.
     */
    public boolean winWait(String title, String text, int timeout) {
        return winVariantBool(title, text, timeout, "WinWait");
    }

    /**
     * Pauses execution of the script until the requested window exists.
     * The script polls for window match every 250 milliseconds or so.
     * @param title The title of the window to check.
     * @return True if success, false otherwise.
     */
    public boolean winWait(String title) {
        return winVariantBool(title, "WinWait");
    }

    /**
     * Pauses execution of the script until the requested window exists.
     * The script polls for window match every 250 milliseconds or so.
     * @param title The title of the window to check.
     * @param text The text of the window to check.
     * @return True if success, false otherwise.
     */
    public boolean winWait(String title, String text) {
        return winVariantBool(title, text, "WinWait");
    }

    /**
     * Pauses execution of the script until the requested window is active.
     * @param title The title of the window to check.
     * @param text The text of the window to check.
     * @param timeout The timeout in seconds.
     * @return True if success, false otherwise.
     */
    public boolean winWaitActive(String title, String text, int timeout) {
        return winVariantBool(title, text, timeout, "WinWaitActive");
    }

    /**
     * Pauses execution of the script until the requested window is active.
     * @param title The title of the window to check.
     * @param text The text of the window to check.
     * @return True if success, false otherwise.
     */
    public boolean winWaitActive(String title, String text) {
        return winVariantBool(title, text, "WinWaitActive");
    }

    /**
     * Pauses execution of the script until the requested window is active.
     * @param title The title of the window to check.
     * @return True if success, false otherwise.
     */
    public boolean winWaitActive(String title) {
        return winVariantBool(title, "WinWaitActive");
    }

    /**
     * Pauses execution of the script until the requested window does not exist.
     * @param title The title of the window to check.
     * @param text The text of the window to check.
     * @param timeout The timeout in seconds.
     * @return True if success, false otherwise.
     */
    public boolean winWaitClose(String title, String text, int timeout) {
        return winVariantBool(title, text, timeout, "WinWaitClose");
    }

    /**
     * Pauses execution of the script until the requested window does not exist.
     * @param title The title of the window to check.
     * @return True if success, false otherwise.
     */
    public boolean winWaitClose(String title) {
        return winVariantBool(title, "WinWaitClose");
    }

    /**
     * Pauses execution of the script until the requested window does not exist.
     * @param title The title of the window to check.
     * @param text The text of the window to check.
     * @return True if success, false otherwise.
     */
    public boolean winWaitClose(String title, String text) {
        return winVariantBool(title, text, "WinWaitClose");
    }

    /**
     * Pauses execution of the script until the requested window is not active.
     * @param title The title of the window to check.
     * @param text The text of the window to check.
     * @param timeout The timeout in seconds.
     * @return True if success, false otherwise.
     */
    public boolean winWaitNoActive(String title, String text, int timeout) {
        return winVariantBool(title, text, timeout, "WinWaitNotActive");
    }

    /**
     * Pauses execution of the script until the requested window is not active.
     * @param title The title of the window to check.
     * @return True if success, false otherwise.
     */
    public boolean winWaitNoActive(String title) {
        return winVariantBool(title, "WinWaitNotActive");
    }

    /**
     * Pauses execution of the script until the requested window is not active.
     * @param title The title of the window to check.
     * @param text The text of the window to check.
     * @return True if success, false otherwise.
     */
    public boolean winWaitNoActive(String title, String text) {
        return winVariantBool(title, text, "WinWaitNotActive");
    }

    private boolean winVariantBool(String title, String text, int timeout, String function) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant vTimeout = new Variant(timeout);
        Variant[] params = new Variant[]{vTitle, vText, vTimeout};
        Variant result = autoItX.invoke(function, params);
        return oneToTrue(result.getInt());
    }

    private boolean winVariantBool(String title, String text, String function) {
        Variant vTitle = new Variant(title);
        Variant vText = new Variant(text);
        Variant[] params = new Variant[]{vTitle, vText};
        Variant result = autoItX.invoke(function, params);
        return oneToTrue(result.getInt());
    }

    private boolean winVariantBool(String title, String function) {
        Variant vTitle = new Variant(title);
        Variant[] params = new Variant[]{vTitle};
        Variant result = autoItX.invoke(function, params);
        return oneToTrue(result.getInt());
    }

    /**
     * Retrieves the text from a standard status bar control.
     * @param title The title of the window to check.
     * @return  The text from a standard status bar control.
     */
    public String statusbarGetText(String title) {
        return autoItX.invoke("StatusbarGetText", title).getString();
    }

    /**
     * Converts the value 1 to true, anything else to false.
     * @param i The value to convert to true/false
     * @return 1 = true, anything else = false.
     */
    private boolean oneToTrue(int i) {
        return i == 1;
    }

    /**
     * Converts the value 1 to true, anything else to false.
     * @param i The value to convert to true/false
     * @return 1 = true, anything else = false.
     */
    private boolean oneToTrue(Variant v) {
        if (v.getvt() == Variant.VariantInt
                || v.getvt() == Variant.VariantShort) {
            return v.getInt() == 1;
        }
        
        return false;
    }
}