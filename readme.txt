PGE 1.1 beta package
--------------------
by Paul Berlin 2003-2004
berlin_paul@hotmail.com
http://pab.dyndns.org

PGE 1.1 is still beta, and I will probably no finish it bacause
I have started on PGE for C++ instead.

PGE is almost done, there are just some bugs and a few missing
features.

If anyone want to finish or enhance PGE you are welcome to.
Send me the final version and I'll put it on my webpage.

If you use PGE or any of this code in your own programs It would
be nice if you gave me credit. You could also send me an e-mail
telling me about it, I would love to see it.

PGE Requires the DirectX 8.1 (or higher) Software Development Kit
to be installed. You can download this from www.microsoft.com.
PGE uses FMOD for sound. FMOD can be found at http://www.fmod.org.

What's in this package
----------------------
This package contains:

* PGE 1.1
* PgeSound
* PgeTiles

PGE 1.1
-------
What needs to be done and known bugs:

* pgeFile
You use pgeFile to save/load data from files.
pgeFile is DONE.

* pgeFont
You use pgeFont to create fonts to draw text with.
pgeFont is DONE.

* pgeKeyboard
You use pgeKeyboard to read input from the keyboard.
pgeKeyboard is DONE.

Suggestions:
1) Rewrite so you use a Poll() function to get the data
from the keyboard (like pgeMouse) once before you check
all wanted keys. Now you do this everytime you check one
key in the KeyDown function.

* pgeMain
You use pgeMain to initialize PGE and render to screen.
pgeMain is DONE.

Suggestions:
1) Add some way to enumerate and select Render devices.
Now the default device is always used.

* pgeMouse
You use pgeMouse to read input from the mouse.
pgeMouse is DONE.

* pgeSound
You use pgeSound to play sounds via FMOD.
1) I have not been able to work out how to load sound into
FMOD from files in an byte array. This needs to be fixed
to let you load Soundsets made with PgeSound.

Suggestions:
1) PGE is using an older version of FMOD... update PGe
to use the newest version.

* pgeSprite
You use pgeSprite to render sprites on screen.
There are probably a few bugs in pgeSprite.

* pgeText
You use pgeText to easy render text on screen.
pgeText is DONE.

* pgeTexture
You use pgeTexture to load textures.
pgeTexture is DONE.

* pgeTileset
You use pgeTileset to load tilesets and create sprites from the tileset.
pgeTileset is DONE.

* pgeTimer
You use pgeTimer to measure time.
pgeTimer is DONE.

PgeSound
--------
You can use PgeSound to create soundsets. Soundsets are a bunch
of sound and music put together in one file.

The problem is that (as seen int the above PGE 1.1 section) I can't
get FMOD to work when loading sounds from a byte array.

What needs to be done in PgeSound:
PgeSound contains some bugs and there are some missing stuff (Like the
About menu option). There's also no documentation.

Other than the above PgeSound works.

PgeTiles
--------
You can use PgeTiles to create tilesets. Tilesets contains
animations, sprites and textures compiled into one file.

This file can be used in PGE with the pgeTileset class to
set up sprites.

What needs to be done in PgeTiles:
Pretty much everything is done. There may be some bugs and you might need
to do some adjustments. The documentation is nowhere done yet.

