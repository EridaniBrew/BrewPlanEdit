Public Class CommandBaseClass
    Protected _comment As String
    Protected _command As String
    Protected _commandDescription As String         ' used to distinguish Comment from flat command, for example
    Protected _commandDocumentation As String       ' provides documentation for the command from ACP User Guide
    Protected _warningFlag As Boolean               ' flags a listItem that should show with red font


    ' Structure containing info for the various # commands
    ' fields are:    key, classType, DefaultString, DocumentationString
    ' Note that documentation string is in html format
    ' Constants used to access 2nd dimension of cmdParms
    Public Const CMDPARMKey = 0
    Public Const CMDPARMType = 1
    Public Const CMDPARMDefaultString = 2
    Public Const CMDPARMDocumentation = 3

    Private Const DocCombo = "Used only when specifying a filter group. For example: <p/><span id=emph>#Count 5,10,5,15</span>" & _
    "<h2 id=emph>#Interval</h2>Set the final target exposure interval(s) for subsequent targets (sec.). For example: <p/><span id=emph>#INTERVAL 31.5</span><p/><span id=emph>#INTERVAL 180,240,180,180</span>" & _
    "<h2 id=emph>#Filter</h2>Required if the system has filters. Set the filter(s) for subsequent targets. If the filter name is not recognized, an error is logged at plan start, and the plan will not run. For example: <p/><span id=emph>#FILTER Blue</span><p/><span id=emph>#FILTER Red,Clear,Green,Blue</span> " & _
    "<h2 id=emph>#Binning</h2>Sets the binning factor(s) for subsequent targets. Note that some detectors don't support arbitrary binning values. Consult the documentation for your detector for specifics. Note: for auto-calibration, masters of the binned size must be available in MaxIm's calibration groups. For example: <p/><span id=emph>#BINNING 4</span><p/><span id=emph>#BINNING 2,1,2,2</span>"

    Public cmdParms(,) As String = {{"#autoguide", "CFlatCommand", "#Autoguide", "Forces the next target's images to be guided, regardless of the setting of ACP's ""enable autoguiding"" preference or the duration of the exposure(s). "},
                            {"#afinterval", "CAfInterval", "#AfInterval 60", "Turns on periodic autofocus and forces an autofocus at the start (or resumption) of the plan. The interval is given in minutes. If an #AUTOFOCUS directive is seen, it overrides a scheduled autofocus, and the time to the next autofocus is reset to the interval. This directive may appear anywhere in the plan, and the value given in the last appearance will be used for the entire plan. For example, to start the plan with an autofocus, then do an autofocus every 30 minutes: <p/><span id=emph>#AFINTERVAL 30</span>"},
                            {"#alwayssolve", "CFlatCommand", "#AlwaysSolve", "Normally, when ACP fails to solve a final/data image in a series (the same target/filter/etc.), it will not try to solve again for that series. This prevents wasting time waiting for plate solves that will probably fail (again). If you want to override this behavior and force ACP to attempt solving every final/data image, include this directive anywhere in your plan. For example: <p/><span id=emph>#ALWAYSSOLVE</span>"},
                            {"#autofocus", "CFlatCommand", "#Autofocus", "Automatically refocus the optical system before each filter group in the filter group for this target. In order to preserve compatibility with the old target-per-filter plan format, this is modified if there is only one filter group. In this case, the autofocus is done once for the target, even if #repeat is greater than one. This requires that FocusMax 3.4.1 or later be installed and autofocus be enabled in ACP's preferences. For example: <p/><span id=emph>#AUTOFOCUS</span>"},
                            {"#bias", "CBias", "#Bias", "Acquire a bias frame using the current target exposure interval. You can use the #REPEAT directive to acquire multiple biases. Multiple biases will be sequence numbered as well as carrying the current #SET number, similar to file naming for light images (except no filter name is included of course). For example: <p/><span id=emph>#BIAS </span><p/>results in one or more files of the form BIAS-Snnn-Rnnn.fts. <p/>An optional complete file path and name may be given, in which case the bias will be created in the given folder with the exact given name. Any existing file with that name will be replaced. For example: <p/><span id=emph>#BIAS D:\MyCalibration\2006012\Bias-Bin2.fts</span> "},
                            {"#binning", "CCombo", "#Binning 1,1,1,1", DocCombo},
                            {"#calibrate", "CFlatCommand", "#Calibrate", "Forces calibration of the images for this target, even if ACP's auto-calibration preference is turned off (it is redundant if ACP's auto-calibration is turned on). This will not cause calibration of pointing exposures, only the final images. For example: <p/><span id=emph>#CALIBRATE</span>"},
                            {"#chain", "CChain", "#Chain myPlan.txt", "When encountered during the last (or only) set, immediately stops reading image acquisition lines from the current plan file, queues a new run of AcquireImages.js with the new plan, then exits. A chained-to plan is thus run in a separate invocation of AcquireImages.js, and starts with conditions identical to those when the same plan is run directly. Use this to chain together plans, each of which might take several sets of images, then wait for a while, then run the new plan which would also take several sets of images. For example: <p/><span id=emph>#CHAIN C:\Program Files (x86)\ACP\Plans\LateNight.txt</span> <p/>or if you just specify a file name, the plan is assumed to be in the same folder as the plan being chained-from. For example: <p/><span id=emph>#CHAIN LateNight.txt</span>    ; In current plan's folder"},
                            {"#chainscript", "CChainScript", "#ChainScript myScript.vbs", "When encountered during the last (or only) set, immediately stops reading image acquisition lines from the current plan file, terminates AcquireImages.js, and starts the given ACP script. If AcquireImages fails or is aborted, the chain will not occur. <p/>The argument is either the full path/name or just the file name only of the script to be chained-to. If only the script file name is given, it is assumed to be in the ACP scripts folder. For example:<p/><span id=emph>#CHAINSCRIPT C:\Program Files (x86)\ACP Obs Control\Scripts\Cleanup.vbs</span><p/><span id=emph>#CHAINSCRIPT Cleanup.vbs </span>   ; In ACP script folder"},
                            {"#chill", "CChill", "#Chill -10", "If needed, turns on the imager's cooler and waits for 5 seconds. In any case, the imager's temperature setpoint is changed to the given temperature (deg. C). After the change, #chill waits for up to 15 minutes for the cooler to reach a temperature within the given tolerance (or 2 degrees, default) of the setpoint. This is actually a type of target, so you can wait before it, have the imager cooled, then wait again so that imaging starts later. If the cooler does not reach the given temperature and tolerance, the plan fails with an error. For example: <p/><span id=emph>#CHILL -35.0</span><p/><span id=emph>#CHILL -32.5, 0.2</span><p/>If your application requires tight temperature tolerances, you can include one of these directives for every target. Thus, before starting on a target, ACP will change or verify the cooler temperature, and fail if it does not meet your criteria."},
                            {"#completionstate", "CCompletionState", "#CompletionState 2,4,1,3,1", "The number of sets, targets in the current set, repeats in the current target, filter groups in the current repeat, and images in the current filter group, that have been completed. Do not include this in your plans, it is automatically inserted in all plans by AcquireImages.js each time a target is completed, then removed if and when the plan runs to completion (at which time #STARTSETNUM is adjusted as described above). Its main use is to allow an interrupted plan to resume at the point where the interruption occurred. For example: <p/><span id=emph>#COMPLETIONSTATE 2,4,1,3,1</span>"},
                            {"#count", "CCombo", "#Count 1,1,1,1|#Interval 60,60,60,60|#Binning 1,1,1,1|#Filter Red,Red,Red,Red", DocCombo},
                            {"#dark", "CDark", "#Dark", "Acquire a dark or bias frame using the current target exposure interval. If you set #INTERVAL to 0 before using #DARK, ACP will acquire a bias frame, and the file naming will be adjusted. It is recommended, however, to use the #BIAS directive described below. You can use the #REPEAT directive to acquire multiple darks or biases. Multiple darks/biases will be sequence numbered as well as carrying the current #SET number, similar to file naming for light images (except no filter name is included of course). For example: <p/><span id=emph>#DARK</span> <p/>results in one or more files of the form Dark-Snnn-Rnnn.fts, or if the preceding #INTERVAL was 0, Bias-Snnn-Rnnn.fts. <p/>An optional complete file path and name may be given, in which case the dark or bias will be created in the given folder with the exact given name. Dark vs bias name changing and sequencing are not done. Any existing file with that name will be replaced. For example: <p/><span id=emph>#DARK D:\MyCalibration\2006012\Dark-Bin2.fts</span>"},
                            {"#dawnflats", "CDawnFlats", "#DawnFlats", "When encountered during the last (or only) set, immediately stops reading image acquisition lines from the current observing plan file, terminates AcquireImages.js, and starts ACP's automatic sky-flat script AutoFlat.vbs. If AcquireImages fails or is aborted, the auto-flats will not occur. See #DUSKFLATS above, and Automatic Flat Frames. Will result in an error is the observatory is configured to use a light panel for flats. <p/>If no argument is supplied, there must be a default flat plan named defaultdawnflat.txt or just defaultflat.txt in the Local User's default plans folder or AcquireImages will not try to start AutoFlat. This avoids AutoFlat stalling waiting for flat plan input. If an argument is supplied it can be either a full path to a flat plan, or just a flat plan file name. If just the flat plan file name is given, it is assumed to be in your default Plans folder. For example:<p/><span id=emph>#DAWNFLATS        ;Need standard flat plan defaultflat.txt in user's default plans folder <p/><span id=emph>#DAWNFLATS 20060122-dawn-flats.txt     ; In user's default plans folder<p/><span id=emph>#DAWNFLATS C:\MasterCalibration\LRGB-Standard-Flats.txt "},
                            {"#defocus", "CDefocus", "#Defocus -150", "Moves the focuser the given number of integer steps away from proper focus just before acquiring each subsequent image. The focus position is restored immediately after acquiring the image, but this directive does carry from target to target, so unless changed, the focus will be moved away from proper focus before each subsequent image. This does not affect pointing images. For example: <p/><span id=emph>#DEFOCUS -150</span>"},
                            {"#domeopen", "CFlatCommand", "#DomeOpen", "Opens the shutter or roll-off roof, and waits until the shutter or roof is actually open. Will un-home or un-park the dome if needed. Effective only during the first or only set-loop of the plan. This is actually a type of target, so you can wait before it, have the shutter or roof opened, then wait again so that imaging starts later. For example: <p/><span id=emph>#DOMEOPEN</span>"},
                            {"#domeclose", "CFlatCommand", "#DomeClose", "Closes the shutter or roll-off roof, and waits until the shutter or roof is actually closed. Effective only during the last or only set-loop of the plan. For example: <p/><span id=emph>#DOMECLOSE</span>"},
                            {"#dither", "CDither", "#Dither 4", "Offset each image in a repeat-set by some small amount away from the original target location. Works for both guided and unguided images. If no parameter is given, ACP uses a value of 5 main imager pixels for dithering (see below). Normally, this value will be appropriate for achieving the noise reduction effect of dithering. Dithering is done by generating two uniform random numbers ranging from minus to plus the ""amount"". One is applied in the X direction, the other in the Y direction. Note that you must supply a value for the guider's plate scale in order for ACP to calculate main imager pixels for guided dithering. If you fail to do this, a warning message will appear in your run log and dithering will be in guider pixels.<p/>If given, the parameter specifies the maximum amount in each axis of this offset in fractional pixels. A parameter value of 0 disables dithering. The random offsets are applied independently in X and Y and are always relative to the initial position. For example: <p/><span id=emph>#DITHER</span>        ; Automatic dithering<p/><span id=emph>#DITHER 3.0</span>  ; 3 pixels dither on the image<p/><span id=emph>#DITHER 0</span>     ; Disable dithering"},
                            {"#dir", "CDir", "#Dir", "Temporarily change the directory into which all subsequent images are to be stored. This can be a relative or full (with a drive letter) directory path, with multiple levels. If relative, the folder is relative to the default image folder as configured in the Local User tab of ACP Preferences (or for web users, their images folder). The folder, including all intermediate levels, is created if needed. For example: <p/><span id=emph>#DIR C:\Special\Comet Search\28-Sep-2003</span>  ;(absolute)<p/><span id=emph>#DIR Photometric Standards\Landolt</span>  ;(relative) <p/>If no folder name is given, this will switch back to the default image folder as configured in the Local User tab of ACP Preferences (or for web users, their images folder) plus the usual date-based subfolder. For example: <p/><span id=emph>#DIR</span>          ; Restore default image folder"},
                            {"#duskflats", "CDuskFlats", "#DuskFlats", "The plan starts by acquiring a series of automatic sky flats at dusk via the AutoFlat.vbs script (which is run under control of AcquireImages.js). See #DAWNFLATS below, and Using Automatic Flat Frames. Will result in an error is the observatory is configured to use a light panel for flats. <p/>If no argument is supplied, there must be a default flat plan named defaultduskflat.txt or just defaultflat.txt in the Local User's default plans folder or AcquireImages will not try to start AutoFlat. This avoids AutoFlat stalling waiting for flat plan input. If an argument is supplied it can be either a full path to a flat plan, or just a flat plan file name. If just the flat plan file name is given, it is assumed to be in your default Plans folder. For example: <p/><span id=emph>#DUSKFLATS</span>        ;Need standard flat plan defaultflat.txt in user's default plans folder <p/><span id=emph>#DUSKFLATS 20060122-dusk-flats.txt</span>     ; In user's default plans folder<p/><span id=emph>#DUSKFLATS C:\MasterCalibration\LRGB-Standard-Flats.txt</span>"},
                            {"#filter", "CCombo", "#Filter Clear,Clear,Clear,Clear", DocCombo},
                            {"#interval", "CCombo", "#Interval 60,60,60,60", DocCombo},
                            {"#manual", "CManual", "#Manual MyTarget", "Acquire an image at the current telescope location. No pointing updates or slews will be done. This is actually a type of target, so don't include a target line. Include an object name. For example: <p/><span id=emph>#MANUAL MyImage</span> <p/>If you don't include an object name, the current date/Time will be used. For example: <p/><span id=emph>#MANUAL </span><p/>results in an image file name of Manual-dd-mm-yyyy@hhmmss-Snnn-Rnnn-filter.fts"},
                            {"#minsettime", "CMinSetTime", "#MinSetTime 00:05", "The minimum amount of time that a set is allowed to take. This can be used to limit the number of sets per unit time. For example: <p/><span id=emph>#MINSETTIME 00:05</span> <p/>will tell ACP to wait until at least 5 minutes has elapsed before starting the next set."},
                            {"#nopointing", "CFlatCommand", "#NoPointing", "Prevent the pointing update prior to the target. Harmless if auto-center is disabled in Preferences. For example: <p/><span id=emph>#NOPOINTING</span> "},
                            {"#nopreview", "CFlatCommand", "#NoPreview", "Prevent the generation of preview images for the web System Status display. This can save significant time per image, maximizing efficiency at the cost of no ""last image preview"" thumbnail or clickable light box image. For example: <p/><span id=emph>#NOPREVIEW</span> "},
                            {"#nosolve", "CFlatCommand", "#NoSolve", "Prevent final/data image plate solving for all of the images of the current target. Harmless if final/data image solving is disabled in Preferences. For example: <p/><span id=emph>#NOSOLVE</span> "},
                            {"#noweather", "CFlatCommand", "#NoWeather", "Disconnects the weather input. This is provided so that you can do calibration frames (darks/biases) in unsafe weather without the weather safety interrupt. Normally follows #domeclose. Weather will not be disconnected if the dome or roof is open. If you have no dome or roof, this will disconnect the weather, so beware! This latter logic is for special cases where the roof or enclosure is not under ACP's control and is sure to be closed in unsafe weather by another means (for example observatory pods housing multiple telescopes). For example: <p/><span id=emph>#NOWEATHER</span>"},
                            {"#pointing", "CFlatCommand", "#Pointing", "Schedule a pointing update prior to the target. This will work even if auto-center is disabled in Preferences. Thus, you can use #POINTING as a means to manually control when pointing updates occur in a plan. For example:<p/><span id=emph>#POINTING</span> "},
                            {"#posang", "CPosAng", "#PosAng 240.5", "Required if a rotator is connected in ACP. If a rotator is installed and connected in ACP, sets the position angle for subsequent images. The value of the position angle ranges from 0 up to but not including 360 degrees. 0 Degrees is pole-up, and the angle increases counterclockwise, that is, north toward east. The rotator will be positioned correctly regardless of GEM meridian flip, and the guider will be adjusted accordingly as well. For example: <p/><span id=emph>#POSANG 240.5</span>"},
                            {"#quitat", "CQuitAt", "#QuitAt 1:00 PM", "Set a ""quitting time"" at which the plan will stop acquiring images. The quitting date/time is in UTC, and is interpreted the same as for #WAITUNTIL. If you specify #DAWNFLATS, #CHAIN, or #CHAINSCRIPT, these actions will still occur after the plan ends. For example: <p/><span id=emph>#QUITAT 7/1/01 08:22</span> <p/>If the plan completes before the quit date/time is reached, it ends as usual. If only a time is given, it will always wait until the given time, even if it was just passed (it will wait till it is that time again)."},
                            {"#trackon", "CFlatCommand", "#TrackOn", "Initiates orbital tracking of solar system bodies. This remains in effect until cancelled by #TRACKOFF. Orbital tracking will not be done except for solar system bodies, so non-solar-system targets may be intermixed without harm. Autoguiding will not be done if orbital tracking is active. Note that orbital tracking requires orbital elements as the target specification (major planet targets will also be tracked). For example: <p/><span id=emph>#TRACKON</span>"},
                            {"#trackoff", "CFlatCommand", "#TrackOff", "Cancels orbital tracking. This remains in effect until re-enabled with #TRACKON. For example: <p/><span id=emph>#TRACKOFF</span>"},
                            {"#readoutmode", "CReadoutMode", "#ReadoutMode Normal", "Selects the imager's readout mode for the current target and all subsequent targets. The imager must support readout modes, and the name you give must be supported by your imager. You can see which readout modes (if any) are supported by looking on the MaxIm DL CCD control window's ""Expose"" tab. Pointing exposures will always use Fast or Normal, so this will not impact pointing update times. For example: <p/><span id=emph>#READOUTMODE 8 MPPS (RBI Flood)</span>"},
                            {"#repeat", "CRepeat", "#Repeat 2", "Tells script to take the given number of filter groups of the next target or dark/bias frame (#DARK) in a row. #REPEAT may be combined with #SETS. For example: <p/><span id=emph>#REPEAT 5</span>"},
                            {"#sets", "CSets", "#Sets 2", "Repeat the entire plan a given number of times. The images are acquired in round-robin order. This directive may appear anywhere in the plan. If it appears more than once, the last value is used for the plan. The default is a single set. For example: <p/><span id=emph>#SETS 3</span>"},
                            {"#screenflats", "CScreenFlats", "#ScreenFlats", "When encountered during the last (or only) set, immediately stops reading image acquisition lines from the current observing plan file, terminates AcquireImages.js, and starts ACP's automatic flat script AutoFlat.vbs. If AcquireImages fails or is aborted, the auto-flats will not occur. See Automatic Flat Frames. Will result in an error is the observatory is configured for sky flats. <p/>If no argument is supplied, there must be a default flat plan named defaultflat.txt in the Local User's default plans folder or AcquireImages will not try to start AutoFlat. This avoids AutoFlat stalling waiting for flat plan input. If an argument is supplied it can be either a full path to a flat plan, or just a flat plan file name. If just the flat plan file name is given, it is assumed to be in your default Plans folder. For example:<p/><span id=emph>#SCREENFLATS </span>       ;Need standard flat plan defaultflat.txt in user's default plans folder <p/><span id=emph>#SCREENFLATS 20060122-flats.txt</span>     ; In user's default plans folder<p/><span id=emph>#SCREENFLATS C:\MasterCalibration\LRGB-Standard-Flats.txt </span>"},
                            {"#shutdown", "CFlatCommand", "#Shutdown", "At the end of the run, parks the scope and shuts down the camera and cooler. If dome control is active, and if the ""Automatically park or home and close AFTER the scope is parked"" option is set, then the dome will be parked or homed and the shutter or roll-off roof will be closed. This may be used with #DAWNFLATS, and shutdown will occur after dawn flats have been taken. For example: <p/><span id=emph>#SHUTDOWN</span> "},
                            {"#shutdownat", "CShutdownAt", "#ShutdownAt 1:00 PM", "Same as #QUITAT, except the scope is parked and the camera is shut down at the quitting time, or at normal exit. The shutdown time is in UTC, and is interpreted the same as for #WAITUNTIL. For example: <p/><span id=emph>#SHUTDOWNAT 7/1/06 08:22</span> <p/>If the plan completes before the shutdown date/time is reached, it acts as though a #SHUTDOWN directive was given instead. If only a time is given, it will always wait until the given time, even if it was just passed."},
                            {"#stack", "CFlatCommand", "#Stack", "Combines repeated images within one filter group without aligning into a single image. Individual images used in the stack are preserved. File names will have -STACK in place of the repeat number. This is most useful when doing orbital tracking. See #TRACKON. The stacked image is saved in IEEE floating-point FITS format to preserve the dynamic range. For example: <p/><span id=emph>#STACK</span>"},
                            {"#stackalign", "CFlatCommand", "#StackAlign", "Combines repeated images within one filter group and aligns images into a single image. Individual images used in the stack are preserved. File names will have STACK in place of the repeat number. Use this for all stare-mode image sets. The stacked image is saved in IEEE floating-point FITS format to preserve the dynamic range. For example:<p/><span id=emph>#STACKALIGN</span>"},
                            {"#startsetnum", "CStartSetNum", "#StartSetNum 2", "The starting set number used in naming image files. Do not include this in your plans, it is automatically inserted in all plans by AcquireImages.js. Each time the plan runs to completion, this number is incremented by the number of sets specified in #SETS or by 1. Its main use is to prevent overwriting of images when the same plan is run multiple times. For example: <p/><span id=emph>#STARTSETNUM 6</span>"},
                            {"#subframe", "CSubFrame", "#Subframe 0.5", "Sets the fraction of the chip to be used for subsequent images. Legal values are 0.1 to 1.0 (full frame). For example, if the chip is 1K by 1K (1024 by 1024), a SUBFRAME of 0.5 will result in using the center 512 by 512 pixels of the chip. For example: <p/><span id=emph>#SUBFRAME 0.5</span>"},
                            {"#tag", "CTag", "#Tag name=value", "Adds a named tag to the target. This directive does not affect the image acquisition process; it simply attaches the tag name and value to the target. You can specify as many of these as you want (each with different names) for any target. The tag name(s) and value(s) will be echoed to the run log, but this is most useful when you have custom actions defined for TargetStart and TargetEnd. These custom actions are passed a Target object as a parameter. Within the custom action, you can refer to tags by their name (as you defined them) with the syntax Target.Tags.name. Thus, you can use tags to alter the action of TargetStart and TargetEnd based on the tags' value(s). This is an expert feature and allows powerful custom logic to be implemented. The syntax is #TAG name=value. There must be an '=' in the #TAG directive. For example: <p/><span id=emph>#TAG type=reference star</span><p/>This will attach a tag 'type' with the value 'reference star' to the target."},
                            {"#waitfor", "CWaitFor", "#WaitFor 30", "Pause for the given number of seconds before processing the next target. For example: <p/><span id=emph>#WAITFOR 30</span>"},
                            {"#waituntil", "CWaitUntil", "#WaitUntil 0,3:30 AM", "<b>(date/time, see below for alternate form)</b><p/>Pause during a specific set (see #SETS) until the given UTC date/time or (only) time. The first parameter is the set number for the pause, the second is the date/time at which to resume. The set number may range from 0 through the number of sets given by the #SETS directive. If there is no #SETS directive on the plan, the set number must be 1. If the set number is 0, it means ""wait on all sets"". This is useful, when only a time is given, for plans that are stopped before completion then resumed on subsequent nights. If a complete date/time is given, and has passed, the directive is ignored. If only a time is given, it will wait for up to 12 hours. If the time is less than 12 hours in the past, it will not wait. The idea is that the time is relative to that observing night, and may be re-used on the next night. See the note below. For example: <p/><span id=emph>#WAITUNTIL 1, 21-Apr-2011 08:02:00</span><p/>Wait until 08:02 UTC only if set #1 and only if 21-Apr-2011<p/><span id=emph>#WAITUNTIL 2, 08:32:00</span><p/>Wait until 08:32 UTC if set #2 no matter what the date is<p/><span id=emph>#WAITUNTIL 0, 09:21:00</span><p/>Wait until 09:21 UTC on any night on any set # (set #0 means ""all sets"")<p/><b>(sun-down angle, see above for alternate form)</b><p/>Pause during a specific set (see #SETS) until the Sun gets below the given angle, degrees (must be a negative number). The first parameter is the set number for the pause, the second is the negative sun-down angle (degrees) at which to resume. The set number may range from 0 through the number of sets given by the #SETS directive. If there is no #SETS directive on the plan, the set number must be 1. If the set number is 0, it means ""wait on all sets"". This is mostly useful for runs that start before dusk. The directive waits for the nearest dusk, so if you start a run with this directive after solar nadir, it will wait until the upcoming dusk! For example: <p/><span id=emph>#WAITUNTIL 1, -10.5</span><p/>Wait until the Sun gets to 10.5 degrees below the horizon if set #1<p/><span id=emph>#WAITUNTIL 0, -10.5</span><p/>Wait until the Sun gets to 10.5 degrees below the horizon on any night on any set # (set #0 means ""all sets"")"},
                            {"#waitzendist", "CWaitZenDist", "#WaitZenDist 40, 30", "Pause until the target is within the given zenith distance (deg) for up to the given time (min). If the target will never get within the given zenith distance, or won't get there within the time limit, it is skipped. A maximum time to wait (minutes) must be included. For Example: <p/><span id=emph>#WAITZENDIST 40, 30 </span><p/>This will wait until the target is within 40 degrees of the zenith for up to 30 minutes."},
                            {"#waitinlimits", "CWaitInLimits", "#WaitInLimits 60", "Pause until the target is within the observatory limits: minimum elevation, horizon, and any tilt-up limit. If target will never meet the criteria, it Is immediately skipped. A maximum time to wait (minutes) must be included. For Example: <p/><span id=emph>#WAITINLIMITS 60</span> <p/>This will wait for the target to rise above the observatory limits for up to 60 minutes."},
                            {"#waitairmass", "CWaitAirMass", "#WaitAirMass 2.5, 30", "Pause until the target is at or below the given air mass. If the target will never get within the given air mass, or won't Get there within the time limit, it is skipped. A maximum time to wait (minutes) must be included. For example: <p/><span id=emph>#WAITAIRMASS 2.5, 30 </span><p/>This will wait until the target is at or below 2.5 air masses for up to 30 minutes."}}

    Private Const HtmlHeader = "<body bgcolor=""#E6E6FA""><font size=2><style>#emph {color: blue; font: italic bold;}</style>"
    Private Const HtmlTrailer = "</font></body>"

    Public Property Comment() As String
        Get
            Return _comment
        End Get
        Set(value As String)
            _comment = value
        End Set
    End Property

    Public Property Command() As String
        Get
            Return _command
        End Get
        Set(value As String)
            _command = value
        End Set
    End Property

    Public Property OrigComment() As String
        Get
            Return _origComment
        End Get
        Set(value As String)
            _origComment = value
        End Set
    End Property

    Public Property Warning() As Boolean
        Get
            Return _warningFlag
        End Get
        Set(value As Boolean)
            _warningFlag = value
        End Set
    End Property

    Public Function UpdateComment(curText As String) As Boolean
        ' if the string has changed, update the chill value
        Dim updated As Boolean = False
        If (curText <> _origComment) Then
            _comment = curText
            _origComment = curText
            updated = True
        End If
        UpdateComment = updated
    End Function


    Public Sub New()
        ' Constructor
        _comment = ""
        _command = ""
        _commandDescription = ""
        _warningFlag = False

    End Sub

    Public Sub Dispose()

    End Sub

    Public Overridable Sub Display()
        ' each 
        MainForm.txtNote.Text = _comment
        ' turn on comment field  for example, CCombo will turn it off
        MainForm.txtNote.Visible = True
        MainForm.lblNote.Visible = True

        MainForm.lblCommandName.Text = _commandDescription
        ' set the documentation
        'MainForm.WebBrowser1.Navigate("about:blank")
        'Dim doc As HtmlDocument = MainForm.WebBrowser1.Document
        'doc.Write(String.Empty)
        MainForm.WebBrowser1.DocumentText = _commandDocumentation
        'MainForm.WebBrowser1.Refresh()
        'Application.DoEvents()

        Dim pnl As Panel
        For Each pnl In MainForm.panelList
            pnl.Visible = False
        Next
        MainForm.pnlBase.Visible = True

    End Sub

    Public Function CleanCommandString(cmd As String) As String()
        ' clean up the command string, extract the command, and retrieve the comment if any
        Dim pieces As String() = {""}
        cmd = RetrieveComment(cmd)
        Dim cleanCmd As String = Trim(cmd.Replace("  ", " "))        ' remove extra spaces
        If (cmd = "") Then
            ' this was a comment line by itself
            _command = ";"
            _commandDescription = "Comment"
            _commandDocumentation = HtmlHeader & "<h2 id=emph>Comment</h2>Do not include the semicolon in the comment; the editor will add it" & HtmlTrailer
        Else
            pieces = cleanCmd.Split(" ")

            If (pieces(0).Substring(0, 1) <> "#") Then
                ' target command
                pieces = cleanCmd.Split(vbTab)
                _command = ""
                _commandDescription = "Target"
                _commandDocumentation = HtmlHeader & "<h2 id=emph>Target</h2>Target is described with either:<p/>" & _
                                        "a) a single object identifier from a catalog (i.e., M 33) <p/>" & _
                                        "<span id=emph>M 33</span><p/>" & _
                                        "b) an identifier followed by the JNow RA/Dec coordinates of the target.<p/>" & _
                                        "<span id=emph>MyM33     20 35 45.3&nbsp;&nbsp;-5 23 50</span><p/>" & _
                                        "The second option allows imaging of a location not identified by a catalog entry, or positioning a known object such as M33 so that guide star is present on the guider</font></body>" & HtmlTrailer
            Else
                ' this is a # command
                _command = pieces(0)
                _commandDescription = pieces(0)
                LookUpDocumentation()
            End If
        End If

        CleanCommandString = pieces
    End Function

    Public Function RetrieveComment(cmd As String) As String
        ' check for a command like #xxx aaa bbb ccc    ' comment here
        ' or    'comment
        ' If found, extract the comment
        ' Note - the quote sign is not included
        _comment = ""
        Dim pos As Integer = cmd.IndexOf(";")
        If (pos > -1) Then
            _comment = cmd.Substring(pos + 1)
            ' remove the comment from the command string
            cmd = cmd.Remove(pos, cmd.Length - pos)
        End If
        Return cmd
    End Function

    Public Function CmdFindCommand(c As String) As Integer
        ' Searches cmdParms for the target command c (something like #chill)
        ' returns -1 if not found, else index into cmdParms
        Dim result = -1
        Dim i As Integer
        c = LCase(c)
        Dim numEntries = cmdParms.GetLength(CMDPARMKey)

        For i = 0 To numEntries - 1
            If (LCase(cmdParms(i, CMDPARMKey)) = c) Then
                Exit For
            End If
        Next
        If (i < numEntries) Then
            result = i
        End If
        CmdFindCommand = result
    End Function

    Private Sub LookUpDocumentation()
        ' if command is #filter / #interval / #binning then use #count for the documentation
        Dim cmd As String = _command
        If ((LCase(_command) = "#filter") Or (LCase(_command) = "#filter") Or (LCase(_command = "#interval"))) Then
            cmd = "#Count"
        End If
        Dim i As Integer = CmdFindCommand(cmd)
        If (i >= 0) Then
            _commandDocumentation = HtmlHeader & "<h2 id=emph>" & cmdParms(i, CMDPARMKey) & "</h2>" & cmdParms(i, CMDPARMDocumentation) & HtmlTrailer
        End If
    End Sub

    ' control updating events
    Private WithEvents myTxtNote As TextBox = MainForm.txtNote
    Protected _origComment As String            ' save value to see if update is required

    Private Sub txtNote_Enter(sender As Object, e As EventArgs) Handles myTxtNote.Enter
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            _origComment = myTxtNote.Text
        End If
    End Sub

    Private Sub txtNote_KeyPress(sender As Object, e As KeyPressEventArgs) Handles myTxtNote.KeyPress
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            If (e.KeyChar = vbCr) Then
                txtNote_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub txtNote_Leave(sender As Object, e As EventArgs) Handles myTxtNote.Leave
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            If (UpdateComment(myTxtNote.Text)) Then
                MainForm.GetPlan().Update()       ' update the plan
            End If
        End If
    End Sub

End Class
