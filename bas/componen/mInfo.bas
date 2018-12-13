Attribute VB_Name = "mInfo"
'==================================================================================================
'vbComCtl.vbp                           7/15/05
'
'           PURPOSE:
'               Replace the functionality of these files:
'
'                   mscomctl.ocx       1042 KB
'                   mscomct2.ocx        633 KB
'                   comct332.ocx        406 KB
'                   richtx32.ocx        200 KB
'                   msmask32.ocx        163 KB
'                   comdlg32.ocx        138 KB
'
'                          total:      2578 KB
'
'               In addition:
'
'                   Versatile custom drawn popup menus in ucPopupMenus.ctl supporting
'                   custom colors, background bitmap with image processing for highlights,
'                   button/gradient/standard highlights, infrequently used items, auto-redisplay
'                   for creation of toolbar-like checked menus, and more.
'
'                   An implementation of the hotkey control.
'
'                   A comctl 6.0 friendly reimplementation of the intrinsic vb frame.
'
'                   A scrolling container.  Place as many controls as you need and it will take on
'                   scrollbars automatically if sized smaller than the constituent controls.
'
'                   Effecient gdi management in mGDI.bas.  Identical fonts/pens/brushes are reused globally and a
'                   reference count is kept.
'
'                   The implementation of an ambient font type enchances the font effeciency even further
'                   than the gdi management alone.  Almost no data needs to be stored in the final
'                   executeable (via the property bag) and controls can easily track changes to the
'                   container's font on the fly.
'
'                   All allocated resources are managed when the bDebug compiler switch is turned on, and
'                   this can alert the programmer to leaks such as gdi object, memory, string, menu and module handles.
'
'                   More advanced functionality is exposed for the ComCtl32.dll controls, including:
'
'                       XP visual styles
'
'                       Simple creation of IE-like rebar/toolbar combinations
'                       supporting chevrons and menu tracking mode
'
'                       Significantly increased speed and lowered memory use; especially
'                       in the large collection based controls, listview and treeview
'
'                       All of the font options of a LOGFONT structure are exposed
'                       by cFont.cls and used by all font-displaying controls.
'
'               Noteable exclusions:
'                   The only control excluded from the files listed above is the FlatScrollBar control.
'
'                   Fewer design time interfaces.  There is a font property
'                   page but other properties are set through the property
'                   window.  Complex property information such as the columns
'                   collection for a listview or a buttons collection for a
'                   toolbar must set at run time.
'
'               Known issues:
'                   ucMonthCalendar incorrectly displays multiple selection when
'                   linked to version 6.
'
'                   ucDateTimePicker dropdown flashes and disappears if the
'                   dropdown arrow is clicked while a textbox is in focus.  It
'                   behaves correctly if a command button is in focus.
'
'                   ucListView columnheaders display a black rectangle where the icon
'                   should be when drag-dropping a column and linked to version 6.
'
'                   When on a system that does not support alpha blending, ImageList_Drag*
'                   for OLE Drag Drop redraws very incorrectly when switching between
'                   applications and flickers when the target control is scrolling.
'
'                   In the VB ide, sometimes the cursor is a cross(as though drawing a control) when you first
'                   open up a usercontrol or form designer window, even though you have not chosen
'                   a control from the toolbox.  When this happens, if you click inside the designer
'                   then the ide crashes. If you get this cursor, you can press escape or close and re-open
'                   the designer to get back to normal.  This seems to usually happen when editing the code of
'                   the usercontrol.  Although it happens more often with a project group containing
'                   vbComCtl, it also also can happen if the client program is referencing the
'                   binary file.
'
'                   Animation control stops playing temporarily if you initiate a window drag operation by clicking
'                   on the form's title bar, drag the whole form below the screen out of view and back up again,
'                   and keep the drag operation in progress by continuing to hold the mouse button.
'
'                   When using the popupmenu's redisplay style for a menu item in a submenu, the hierarchy
'                   does not redisplay when the user double-clicks.
'
'           DEPENDENCIES:
'               vbComCtl.tlb
'               Interfaces.tlb
'               stdole2.tlb
'
'           COPYRIGHTS:
'               I won't sue you no matter how you use it, provided that
'               you won't sue me no matter how it works or fails to work.
'
'==================================================================================================
