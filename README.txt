HI!

About this stuff.

The very 1st MUME Online Map development started on 5th of April 2003, because I got "lost" near Tharbad and almost died! And I had been thinking of doing this for some time (it seemed quite an easy challenge and it was).
Prototype was built in 2 days and "friendly" version was out with 2 weeks. Half of the time I f***ng with bits and pixels.
This was my first Visual Basic and desktop program, but thanks sockets library (MSWINSCK.ocx) it was pure fun.
There is no special architecture, just code that works, fast as hell and readable even for a baby alien.
I think I gave quite a push to other developers out there, but lets be honest - there is only #1 MOM (coined).
I left the code as is, with all the GODMODE features, map decryption passwords and legacy stuff - HAVE FUN!


/old_versions/prototype
---------------------------------------
Proof of concept map editor and "database" was in Excel file (NOC.xls). Of course this was slow as hell, thus the later versions (hah)!
Document "worldmap.doc" was to understand the maximum rooms in the world (do I have enough memory) that also gave me insight how to build the world data matrix.
Document "bitmapping.xls" was mapping all room variables, to seek and check what to draw.
Also you can treat yourself with "Logs/output.log" that I found accidentally, this describes the emotions with very first tests and of course my mature personality.


/old_versions/beta
---------------------------------------
I'm pretty sure I spent a week editing the bitmaps to make different terrains blend as fluently as possible. 
This time I went all in with MS fastest Access Database "world.mdb". The map was loaded and decrypted in only THREE minutes (fkghell).
After not understanding why everything was so slow, I just made my own CSV text file and structure, result was full world load in 2 seconds (22454 rooms and relations).
By making an unique hash from room description + attributes/exits, the sync lookup was instant and only east of bree with duplicate rooms had issues.
Code is based on bitwise checks, so the program would be usable even on the Intel Pentium on Windows 98 (for my dear friends).


/global
---------------------------------------
MOM core source code
It was a nice challenge to build the world mapping editor UX, with smart portals and lines.
The easiest was to make world map view, I think that took an hour.


/MUME
---------------------------------------
"GODMODE" version - MOM UI source code (fools, friends and me)
Included a full map and advanced features:
#1 feature was "where" also showed player initials on map rooms
#2 feature was "flee-but-never-lost" algorithm, that 99.99% was never lost, also the sync time was instant. This was the part all other mappers failed, until gods helped the mortals with XML output. I didn't understand how come people couldn't do basic string parsing (CR/LF and logic stuff).
#3 feature was "undo", where after fleeing a "button" took you back to previous room, thus hit-flee was a charm (before the XML)
#4 feature was pre-drawn walking path on map, so you knew when you had mis-spammed, or when to open door
#5 feature was "blind mode", where map continued working after blinded/fog, and did re-sync the moment you could see again
I had it all at 2003 :D

Some code that never got ready enough (it was either code or kill pukes -> result was Warlord):
* magic-timers that know and show all spells uptime
* living meta info on map, where I saw magic casted, encountered enemy, saw tracks, covered in darkness
* never finished group members tab "labinator" with their status, health, effects (every member reported silently to all)

Ideas:
* I was going to visualize MUME as WAR TV with avatars and simple animations, that would run Live and from Logs, but my ego was already saturated


/MUMEPublic
---------------------------------------
Public version - MOM UI source code
Included Arda map with major cities and roads.
   If Not GODMODE Then
      frmMap.mnuMovement.Checked = False: frmMap.mnuMovement.Enabled = False: frmMap.mnuMovement.Visible = False
      frmMap.mnuPlayers.Checked = False: frmMap.mnuPlayers.Enabled = False: frmMap.mnuPlayers.Visible = False
      frmMap.mnuEnemies.Checked = False: frmMap.mnuEnemies.Enabled = False: frmMap.mnuEnemies.Visible = False
      frmMap.mnuTarget.Checked = False: frmMap.mnuTarget.Enabled = False: frmMap.mnuTarget.Visible = False
      frmMap.mnuHere.Checked = False: frmMap.mnuHere.Enabled = False: frmMap.mnuHere.Visible = False
      frmMap.mnuWalk.Checked = False: frmMap.mnuWalk.Enabled = False: frmMap.mnuWalk.Visible = False
      frmMap.mnuReceiver.Checked = False: frmMap.mnuReceiver.Enabled = False: frmMap.mnuReceiver.Visible = False
      frmMap.mnuInformer.Checked = False: frmMap.mnuInformer.Enabled = False: frmMap.mnuInformer.Visible = False
   End If


/SETUP
---------------------------------------
MOM setup project source code.
One thing is to write a program, another thing is to make it work on every Windows version.


/drivers
---------------------------------------
Ah, yes, well, Microsoft and stuff


/tools_graphics
---------------------------------------
An unknown future friend made designs that I received through a friend because he was bored - TY!


/homepage
---------------------------------------
My design of MUME homepage year 2000, with herblores and characters.



C U OUT THERE!

#1 Warlord from Estonia,
Diamonium (.ee) AKA Jaanus Lang
jaanus.lang@gmail.com
