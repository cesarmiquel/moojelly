**Q: I get "Block-map to large. Increase of overall horizontal detail. will probably generate corrupted ROM."**

**A:** due to how SML2 is coded it has a limit to how much detail in the horizontal plain in one block of levels, when it is patching check which levels are grouped and if you want to add more detail to one level take away detail from another in that group.

by horizontal detail I mean:
  * 'XOXOXOX' has 6 changes of block and doesn't compress
  * 'XOOOXXX' has 2 changes of block and compresses to 5 bytes

**Q: why does it mess up the ROM if I put more sprites in a level?**

**A:** there is a global total amount of sprites and items, take some out of another level.

**Q: how do I set the alignment of the pattern tool?**

**A:** press control as you click.