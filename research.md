# Introduction #

This document is not for users, but for people wanting to make there own SML2 hacks. As it goes into detail on file format.

**Header:**

says where to load everything for each level.

0x13 bytes
D4 01 22 00 C0 01 50 00 00 58 00 00 40 09 01 E4 D0 38 00 04

  * Bytes 00 to 03 (D4 01 22 00)	- Mario's starting coordinates. D4 01 = X, 22 00 = Y.
  * Bytes 04 to 07 (C0 01 50 00)	- Screen focus coordinates. C0 01 = X, 50 00 = Y.
  * Byte 08 (00)					- Screen focus shift. Not practically useful in any way.
  * Byte 09 (58)					- No idea. If you figure this one out, be sure to contact me.
  * Byte 0A (00)					- Level number.
  * Bytes 0B and 0C (00 40)			- Tilemap pointers. All from bank 8.
  * Byte 0D (09)					- Map bank.
  * Byte 0E (01)					- Music.
  * Bytes 0F to 11 (E4 D0 38)		- Palettes. E4 = Background, D0 = Sprite, 38 = Sprite.
  * Byte 12 (00)					- Second map byte.
  * Byte 13 (04)					- Time. (Multiply by 0x64 to get the time in Decimal)

**level borders:**

```
Byte 0x0D of the level's header is the Map Bank byte. The Intro level's map bank is 09 (0x5618). 
Byte 0x12 of the level's header is the second map byte. The Intro level's second map is 00 (0x561D). 

To get to the level scrolls array: 0x4000 x Map_Bank. 
- 0x4000 is one Game Boy 'bank' 
- map_bank in this case is 0x09. 
- The result is 0x24000 

When you've gone to the offset, you'll find two pointers; 0x04 and 0x40. 
-The 0x04 points to the start of the scroll box array. 
-0x40 is covered elsewhere. 

Then: (0x24000 + 0x04) + (second_map x 0x30) 
-0x24000 is the result to the last part. 
-Result 0x24004. which is the scroll box array of the intro level 
-All scrollbox arrays are 0x30 bytes long. 
-There are 16 columns and 3 rows of scrollboxes. 
-One scrollbox is 16 blocks high and wide. 

06 04 04 04 04 04 04 04 04 04 04 04 04 04 04 05 
0E 0C 0C 0C 0C 0C 0C 0C 0C 0C 0C 0C 0C 0C 0C 0D 
0A 0C 0D 00 0C 0C 0C 0C 00 00 00 00 00 0F 00 00 

Disregard the first nibble of each byte, as it serves no purpose and is always 0. 
The second nibble is what matters, though. 
Look at the bits of the first scrollbox.6 = 0110 in binary. 
Now think of 0110 as BTLR. B means Bottom, T means Top, L means Left, and R means Right, respectively. 
If a bit is set, it means you cannot see nor go to the screen in the direction the bit represents. 

1101 would mean you had no choice but to go left. 

-Original information from CoolToby, Compiled by RacoonSam, Rewritten and corrected by SmellyMoo 
```

**Tile-map**

```
Notes: Blank sprite means blank sprite. You can add anything to it.

	??? means that I haven't identified the sprite.
	If I've put a tilde before the address, it means that the data is loaded from an entirely different location.
	Everything else is just after the last address.

Tile structure:

	Byte 1 - X
	Byte 2 - Y
	Byte 3 - Tile
	Byte 4 - Properties
				Bit 0 - If set, invert colors
				Bit 1 - Flip horizontally
				Bit 2 - Flip vertically
				Bit 3 - Transparency modifier
		Second nibble of Byte 4 has no effect whatsoever.

Main Map:

	Graphics loaded from 0x34000
	 0x3CC73 - 72 bytes - 18 tiles:	Hippopotamus statue
	 0x3CCBF - 16 bytes -  4 tiles:	Clear-flag, frame 1
	 0x3CCD3 - 16 bytes -  4 tiles:	Clear-flag, frame 2
	 0x3CCE7 - 16 bytes -  4 tiles:	Clear-flag, frame 3
	 0x3CCFB - 36 bytes -  9 tiles:	Turtle head, frame 1
	 0x3CD23 - 36 bytes -  9 tiles:	Turtle head, frame 2
	 0x3CD4B - 36 bytes -  9 tiles:	Turtle head, frame 3
	 0x3CD73 - 36 bytes -  9 tiles:	Turtle head, frame 4
	 0x3CD9B - 32 bytes -  8 tiles:	Wandering cloud
	 0x3CDBF - 16 bytes -  4 tiles:	Thunder
	 0x3CDD3 -  8 bytes -  2 tiles:	Turtle's neck
	 0x3CDDF - 16 bytes -  4 tiles:	???
	 0x3CDF3 - 16 bytes -  4 tiles:	???
	 0x3CE07 - 16 bytes -  4 tiles:	???
	 0x3CE1B - 16 bytes -  4 tiles:	???
	 0x3CE2F -  4 bytes -  1 tile:	Blank sprite

Tree Zone:

	Graphics loaded from 0x39800
	 0x6245E - 24 bytes -  6 tiles:	Crow, frame 1
	 0x6247A - 28 bytes -  7 tiles:	Crow, frame 2
	 0x6249A - 24 bytes -  6 tiles:	Crow, frame 3
	 0x624B6 - 28 bytes -  7 tiles:	Crow, frame 4
	 0x624D6 - 24 bytes -  6 tiles:	Bee, frame 1
	 0x624F2 - 24 bytes -  6 tiles:	Bee, frame 2
	 0x6250E -  4 bytes -  1 tile:	Blank sprite
	 0x62516 -  4 bytes -  1 tile:	Blank sprite
	 0x6251E -  4 bytes -  1 tile:	Blank sprite
	 0x62526 -  4 bytes -  1 tile:	Blank sprite
	 0x6252E -  4 bytes -  1 tile:	???
	 0x62536 -  4 bytes -  1 tile:	???
	 0x6253E -  4 bytes -  1 tile:	???
	 0x62546 -  4 bytes -  1 tile:	???
	 0x6254E -  4 bytes -  1 tile:	Level cleared-ring (Widely-used sprite! If you change this, it will affect all the other rings!)
	 0x62556 -  8 bytes -  2 tiles:	???
	~0x62B7E - 28 bytes -  7 tiles:	Ant, frame 1
	~0x62B9E - 28 bytes -  7 tiles:	Ant, frame 2
	~0x62BBE - 28 bytes -  7 tiles:	Ant, frame 3
	~0x62BDE - 28 bytes -  7 tiles:	Ant, frame 4
	~0x62BFE - 28 bytes -  7 tiles:	Ant, frame 5
	~0x62C1E - 28 bytes -  7 tiles:	Ant, frame 6
	~0x62C3E - 28 bytes -  7 tiles:	Ant, frame 7	
	~0x62C5E - 28 bytes -  7 tiles:	Ant, frame 8

Pumpkin Zone:

	Graphics loaded from 0x38000
	 0x62562 -  8 bytes -  2 tiles:	Skull's glowing eyes, frame 1
	 0x6256E -  8 bytes -  2 tiles:	Skull's glowing eyes, frame 2
	 0x6257A - 12 bytes -  3 tiles:	Ghost flare, frame 1
	 0x6258A - 12 bytes -  3 tiles:	Ghost flare, frame 2
	 0x6259A - 12 bytes -  3 tiles:	Ghost flare, frame 3
	 0x625AA - 12 bytes -  3 tiles:	Ghost flare, frame 4
	 0x625BA - 48 bytes - 12 tiles:	Flying witch, frame 1
	 0x625EE - 48 bytes - 12 tiles:	Flying witch, frame 2
	 0x62622 - 48 bytes - 12 tiles:	Flying witch, frame 3
	 0x62656 - 48 bytes - 12 tiles:	Flying witch, frame 4
	 0x6268A - 48 bytes - 12 tiles:	Flying witch, frame 5
	 0x626BE - 48 bytes - 12 tiles:	Flying witch, frame 6

Mario Zone:

	Graphics loaded from 0x44000
	 0x626F2 -  8 bytes -  2 tiles:	Left eyebrow
	 0x626FE -  8 bytes -  2 tiles:	Right eyebrow
	 0x6270A - 24 bytes -  6 tiles:	Ear, frame 1
	 0x62726 - 16 bytes -  4 tiles:	Ear, frame 2
	 0x6273A -  8 bytes -  2 tiles:	Ear, frame 3
	 0x62746 - 16 bytes -  4 tiles:	Ear, frame 4
	 0x6275A - 24 bytes -  6 tiles:	Ear, frame 5
	 0x62776 - 24 bytes -  6 tiles:	M-symbol on the hat
	 0x62792 -  8 bytes -  2 tiles:	Eye
	 0x6279E - 32 bytes -  8 tiles:	Shoe

Turtle Zone:

	Graphics loaded from 0x39800
	 0x627C2 -  4 bytes -  1 tile:	ZZZ, frame 1 (Invisible)
	 0x627CA -  4 bytes -  1 tile:	ZZZ, frame 2
	 0x627D2 -  8 bytes -  2 tiles:	ZZZ, frame 3
	 0x627DE - 12 bytes -  3 tiles:	ZZZ, frame 4
	 0x627EE - 12 bytes -  3 tiles:	Seaweed, frame 1
	 0x627FE - 12 bytes -  3 tiles:	Seaweed, frame 2
	 0x6280E - 12 bytes -  3 tiles:	Seaweed, frame 3
	 0x6281E - 12 bytes -  3 tiles:	Seaweed, frame 4
	 0x6282E -  4 bytes -  1 tile:	Bubbles, frame 1
	 0x62836 -  4 bytes -  1 tile:	Bubbles, frame 2
	 0x6283E -  4 bytes -  1 tile:	Bubbles, frame 3
	 0x62846 -  4 bytes -  1 tile:	Bubbles, frame 4
	 0x6284E - 24 bytes -  6 tiles:	Group of fishes
	 0x6286A - 24 bytes -  6 tiles:	Group of fishes, flipped

Wario's Castle:

	Graphics loaded from 0x45800
	 0x62886 - 32 bytes -  8 tiles:	Tower tops
	 0x628AA - 20 bytes -  5 tiles:	Flag, frame 1  	 0x628C2 - 20 bytes -  5 tiles:	Flag, frame 2  } - The first byte is the flagpole knob, the other 4 are the flag itself.
 	 0x628DA - 20 bytes -  5 tiles:	Flag, frame 3 /
	 0x628F2 - 36 bytes -  9 tiles:	Lightning
	 0x6291A - 16 bytes -  4 tiles:	Wario, frame 1
	 0x6292E - 16 bytes -  4 tiles:	Wario, frame 2
	 0x62942 - 16 bytes -  4 tiles:	Wario, frame 3
	 0x62956 - 16 bytes -  4 tiles:	Wario, frame 4
	 0x6296A - 16 bytes -  4 tiles:	Wario, frame 5
	 0x6297E - 16 bytes -  4 tiles:	Wario, frame 6
	 0x62992 - 16 bytes -  4 tiles:	Wario, frame 7

	 0x629A6 - 16 bytes -  4 tiles:	Wario, frame 8
	 0x629BA - 16 bytes -  4 tiles:	Wario, frame 9

Minigame hill:

	Graphics loaded from 0x38000
	 0x629CE - 32 bytes -  8 tiles:	Clouds in the background
	~0x62C7E - 40 bytes - 10 tiles:	Coins border

Space Zone:

	Graphics loaded from 0x40000
	 0x629F2 -  4 bytes -  1 tile:	Twinkling star, frame 1
	 0x629FA -  4 bytes -  1 tile:	Twinkling star, frame 1
	 0x62A02 -  4 bytes -  1 tile:	Twinkling star, frame 1
	 0x62A0A - 12 bytes -  3 tiles:	Shooting star, frame 1
	 0x62A1A - 12 bytes -  3 tiles:	Shooting star, frame 2
	 0x62A2A - 12 bytes -  3 tiles:	Shooting star, frame 3
	 0x62A3A -  4 bytes -  1 tile:	Star's eyes, frame 1
	 0x62A42 -  4 bytes -  1 tile:	Star's eyes, frame 2
	 0x62A4A -  4 bytes -  1 tile:	Star's eyes, frame 3
```

**Main-map Path format:**

```
The over-world path pointer array starts at 0x61602 and ends at 0x61A3F. From there, you start reading 8-byte 'crosses'. They're called crosses because they store the info of all the four points of the compass.

Below are the nine very first 'crosses'.

61602 | FF FF | FF FF | FF FF | FF FF |

6160A | FF FF | FF FF | FF FF | FF FF |

61612 | FF FF | FF FF | FF FF | FF FF |

6161A | FF FF | FF FF | FF FF | FF FF |

61622 | FF FF | FF FF | FF FF | FF FF |

6162A | FF FF | FF FF | FF FF | FF FF |

61632 | FF FF | FF FF | FF FF | FF FF |

6163A | FF FF | FF FF | FF FF | FF FF |

61642 | A1 5A | AE 5A | B8 5A | FF FF |

xxxxx | Right | Left | Up | Down |

Now, observe 0x61642 and its first and second byte, but forget the first nibble of byte 2. That gives us A1 xA. Swap the bytes, so you have xA A1. That is our pointer (still disregarding the 5). Then you add A A1 to 0x61000. That's 0x61AA1.

	What does this all mean then?
	It means that upon pushing RIGHT on the map point the cross is in, the game will read the path instructions from 0x61AA1.

Let's try again with the Left; AE xA A AE 0x61AAE

		It works, but certainly isn't the correct way to do it.
		I have no idea what the first nibble of the second byte
		has to do with anything, so if you ever figure it out,
		be sure to contact me.

Now let's actually go to 0x61AA1 (The path instructions for pushing right on cross 9).

01 01 08 08 08 01 01 01 01 01 01 00

These are the instructions. They are like small, one-byte commands to Mario of where to go and how. The first nibble tells you HOW to move, and the second WHERE.

	First nibble:
		0 = Walk
		1 = Half-walk
		2 = Tiny
		3 = Tiny (also)
		4 = Invisible
		5 = Invisible
		6 = Invisible
		7 = Invisible
		8 = Climb
		9 = Walk
		A = Walk (choppy)
		B = Tiny
		C = Invisible
		D = Invisible
		E = Invisible
		F = Invisible half-walk
			Now there's a lot of invisible stuff there,
			you might start to think that they're
			'invisible climbing' or something like that,
			but that would mean that one bit means a property,
			but it doesn't. It might look like it's like that,
			but no, they're not bitwise.

	Second nibble (this one is bitwise):
		Bit 0 = Right
		Bit 1 = Left
		Bit 2 = Up
		Bit 3 = Down
			-Your crazy combinations such as Up-Down and Right-Left are in favor of Up and Right.
			-Diagonal movement very possible.
			-No bits set means Mario being stationary. Can't think of any situation where to use this.

So 01 01 08 08 08 01 01 01 01 01 01 00 would mean the following; Walk right two tiles -> walk down three tiles -> walk right six tiles -> stop.

That is the first real cross in the array, the cross that's left leads to Tree Zone. 
```