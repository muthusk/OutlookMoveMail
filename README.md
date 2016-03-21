# OutlookMoveMail
VBA to Move Mails from one folder to another folder automatically based on timer events
This works on 64 Bit Office / Outlook Pro Plus 2016 - Not tested on other platforms. 
Timer related declarations would change for 32 bit editions.

This script should be loaded as VBA project and Outlook security settings for macros needs to be relaxed unless you are going to sign the macros digitally. 

Most of the code is already available in blogs and articles. I used those to customize for my case. 

My use case was to move mails by a timer event that triggers every 5 minutes to have Lync conversations saved ina particualr folder to be  moved to another folder. Why ? Lync Conversation folder had a different policy applied and could not be changed at outlook client end in my environment.

To do, Document how to sign VBA macros. 

Code is almost self sufficient for some to read and modify.
