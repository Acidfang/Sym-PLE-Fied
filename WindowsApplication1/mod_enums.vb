Public Module mod_enums
    ' from pinvoke.net
    Public Enum WindowShowStyle As UInteger

        ''' <summary>Hides the window and activates another window.</summary>
        ''' <remarks>See SW_HIDE</remarks>
        Hide = 0
        '''<summary>Activates and displays a window. If the window is minimized 
        ''' or maximized, the system restores it to its original size and 
        ''' position. An application should specify this flag when displaying 
        ''' the window for the first time.</summary>
        ''' <remarks>See SW_SHOWNORMAL</remarks>
        ShowNormal = 1
        ''' <summary>Activates the window and displays it as a minimized window.</summary>
        ''' <remarks>See SW_SHOWMINIMIZED</remarks>
        ShowMinimized = 2
        ''' <summary>Activates the window and displays it as a maximized window.</summary>
        ''' <remarks>See SW_SHOWMAXIMIZED</remarks>
        ShowMaximized = 3
        ''' <summary>Maximizes the specified window.</summary>
        ''' <remarks>See SW_MAXIMIZE</remarks>
        Maximize = 3
        ''' <summary>Displays a window in its most recent size and position. 
        ''' This value is similar to "ShowNormal", except the window is not 
        ''' actived.</summary>
        ''' <remarks>See SW_SHOWNOACTIVATE</remarks>
        ShowNormalNoActivate = 4
        ''' <summary>Activates the window and displays it in its current size 
        ''' and position.</summary>
        ''' <remarks>See SW_SHOW</remarks>
        Show = 5
        ''' <summary>Minimizes the specified window and activates the next 
        ''' top-level window in the Z order.</summary>
        ''' <remarks>See SW_MINIMIZE</remarks>
        Minimize = 6
        '''   <summary>Displays the window as a minimized window. This value is 
        '''   similar to "ShowMinimized", except the window is not activated.</summary>
        ''' <remarks>See SW_SHOWMINNOACTIVE</remarks>
        ShowMinNoActivate = 7
        ''' <summary>Displays the window in its current size and position. This 
        ''' value is similar to "Show", except the window is not activated.</summary>
        ''' <remarks>See SW_SHOWNA</remarks>
        ShowNoActivate = 8
        ''' <summary>Activates and displays the window. If the window is 
        ''' minimized or maximized, the system restores it to its original size 
        ''' and position. An application should specify this flag when restoring 
        ''' a minimized window.</summary>
        ''' <remarks>See SW_RESTORE</remarks>
        Restore = 9
        ''' <summary>Sets the show state based on the SW_ value specified in the 
        ''' STARTUPINFO structure passed to the CreateProcess function by the 
        ''' program that started the application.</summary>
        ''' <remarks>See SW_SHOWDEFAULT</remarks>
        ShowDefault = 10
        ''' <summary>Windows 2000/XP: Minimizes a window, even if the thread 
        ''' that owns the window is hung. This flag should only be used when 
        ''' minimizing windows from a different thread.</summary>
        ''' <remarks>See SW_FORCEMINIMIZE</remarks>
        ForceMinimized = 11

    End Enum
    ''' <summary>The set of valid MapTypes used in MapVirtualKey
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum MapVirtualKeyMapTypes As UInt32
        ''' <summary>uCode is a virtual-key code and is translated into a scan code.
        ''' If it is a virtual-key code that does not distinguish between left- and 
        ''' right-hand keys, the left-hand scan code is returned. 
        ''' If there is no translation, the function returns 0.
        ''' </summary>
        ''' <remarks></remarks>
        MAPVK_VK_TO_VSC = &H0

        ''' <summary>uCode is a scan code and is translated into a virtual-key code that
        ''' does not distinguish between left- and right-hand keys. If there is no 
        ''' translation, the function returns 0.
        ''' </summary>
        ''' <remarks></remarks>
        MAPVK_VSC_TO_VK = &H1

        ''' <summary>uCode is a virtual-key code and is translated into an unshifted
        ''' character value in the low-order word of the return value. Dead keys (diacritics)
        ''' are indicated by setting the top bit of the return value. If there is no
        ''' translation, the function returns 0.
        ''' </summary>
        ''' <remarks></remarks>
        MAPVK_VK_TO_CHAR = &H2

        ''' <summary>Windows NT/2000/XP: uCode is a scan code and is translated into a
        ''' virtual-key code that distinguishes between left- and right-hand keys. If
        ''' there is no translation, the function returns 0.
        ''' </summary>
        ''' <remarks></remarks>
        MAPVK_VSC_TO_VK_EX = &H3

        ''' <summary>Not currently documented
        ''' </summary>
        ''' <remarks></remarks>
        MAPVK_VK_TO_VSC_EX = &H4
    End Enum
    <Flags> _
    Public Enum SetWindowPosFlags As UInteger
        ''' <summary>If the calling thread and the thread that owns the window are attached to different input queues, 
        ''' the system posts the request to the thread that owns the window. This prevents the calling thread from 
        ''' blocking its execution while other threads process the request.</summary>
        ''' <remarks>SWP_ASYNCWINDOWPOS</remarks>
        SynchronousWindowPosition = &H4000
        ''' <summary>Prevents generation of the WM_SYNCPAINT message.</summary>
        ''' <remarks>SWP_DEFERERASE</remarks>
        DeferErase = &H2000
        ''' <summary>Draws a frame (defined in the window's class description) around the window.</summary>
        ''' <remarks>SWP_DRAWFRAME</remarks>
        DrawFrame = &H20
        ''' <summary>Applies new frame styles set using the SetWindowLong function. Sends a WM_NCCALCSIZE message to 
        ''' the window, even if the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE 
        ''' is sent only when the window's size is being changed.</summary>
        ''' <remarks>SWP_FRAMECHANGED</remarks>
        FrameChanged = &H20
        ''' <summary>Hides the window.</summary>
        ''' <remarks>SWP_HIDEWINDOW</remarks>
        HideWindow = &H80
        ''' <summary>Does not activate the window. If this flag is not set, the window is activated and moved to the 
        ''' top of either the topmost or non-topmost group (depending on the setting of the hWndInsertAfter 
        ''' parameter).</summary>
        ''' <remarks>SWP_NOACTIVATE</remarks>
        DoNotActivate = &H10
        ''' <summary>Discards the entire contents of the client area. If this flag is not specified, the valid 
        ''' contents of the client area are saved and copied back into the client area after the window is sized or 
        ''' repositioned.</summary>
        ''' <remarks>SWP_NOCOPYBITS</remarks>
        DoNotCopyBits = &H100
        ''' <summary>Retains the current position (ignores X and Y parameters).</summary>
        ''' <remarks>SWP_NOMOVE</remarks>
        IgnoreMove = &H2
        ''' <summary>Does not change the owner window's position in the Z order.</summary>
        ''' <remarks>SWP_NOOWNERZORDER</remarks>
        DoNotChangeOwnerZOrder = &H200
        ''' <summary>Does not redraw changes. If this flag is set, no repainting of any kind occurs. This applies to 
        ''' the client area, the nonclient area (including the title bar and scroll bars), and any part of the parent 
        ''' window uncovered as a result of the window being moved. When this flag is set, the application must 
        ''' explicitly invalidate or redraw any parts of the window and parent window that need redrawing.</summary>
        ''' <remarks>SWP_NOREDRAW</remarks>
        DoNotRedraw = &H8
        ''' <summary>Same as the SWP_NOOWNERZORDER flag.</summary>
        ''' <remarks>SWP_NOREPOSITION</remarks>
        DoNotReposition = &H200
        ''' <summary>Prevents the window from receiving the WM_WINDOWPOSCHANGING message.</summary>
        ''' <remarks>SWP_NOSENDCHANGING</remarks>
        DoNotSendChangingEvent = &H400
        ''' <summary>Retains the current size (ignores the cx and cy parameters).</summary>
        ''' <remarks>SWP_NOSIZE</remarks>
        IgnoreResize = &H1
        ''' <summary>Retains the current Z order (ignores the hWndInsertAfter parameter).</summary>
        ''' <remarks>SWP_NOZORDER</remarks>
        IgnoreZOrder = &H4
        ''' <summary>Displays the window.</summary>
        ''' <remarks>SWP_SHOWWINDOW</remarks>
        ShowWindow = &H40
    End Enum
    Public Enum material_type
        barrier
        bag
        film
        rollbag
        laminate
    End Enum
    Public Enum Print_style
        plain
        line
        screen
        process
    End Enum
    Public Enum job_status
        New_Job = -1
        Pending = 0
        Mounted = 1
        Running = 2
        Finished = 10
    End Enum
End Module
