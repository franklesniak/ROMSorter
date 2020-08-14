# Emulator Background Information

If you are new to the world of arcade emulation, or not familiar with the differences between the emulators (or if you are asking yourself: _what is an emulator?_), it may be helpful for you to review the following background information before you dive into trying to create a ROM list.

## Emulators? ROM Packages? ROMs? Time for a Terminology Check

Many arcade systems were created at a time when microprocessors and computer equipment were very different than today's computers.
As a result, we cannot merely take the arcade game software and run it on a modern microprocessor.
Instead, modern computer systems need to _pretend_ or _simulate_ that they are running the same microprocessor and hardware that the original arcade system had.
Emulation is the process of _pretending_ or _simulating_ a different hardware platform, such as those found in 1980s and 1990s arcade systems.

Once we can simulate an arcade system's original hardware, we still need the game's software.
The community calls the software required to run an arcade game a _read-only memory (ROM) package_, or sometimes just a _ROM_.
The term ROM package reflects the fact that the oldest arcade systems' software was permanently "etched" into read-only memory chips.
And because an arcade system typically had several of these memory chips, we call the entire package of software for one game a _ROM package_–that is to say that you have packaged up the ROMs, and they are ready to play.

## Emulators Relevant to Playing Arcade Games on a Raspberry Pi 4 B

At the time of writing, three arcade emulators work well on the Raspberry Pi 4.
Each working ROM package matches best with one of these three emulators, and so it is very likely that we need to use all three emulators on a Raspberry Pi arcade build.
Therefore, our analysis must consider all three.

The three emulators are:

+ MAME 2003 Plus
+ MAME 2010
+ FBNeo

## Two Versions of MAME

_MAME_ is an abbreviation for Multiple Arcade Machine Emulator.
It is still under active development, releasing a few versions per year, each adding support for new ROM packages or improving the support of existing ones.
Historically, the community at large called a milestone release in a given year "MAME YYYY," where "YYYY" is the release year.
For example, the community knows MAME version 0.78 as MAME 2003 because the development team released it in 2003.
The importance of this naming becomes clear in a moment.

The MAME team wrote the emulation software for modern, full-powered x86 computers and not for lower-performance ARM systems like the Raspberry Pi.
Luckily, some savvy developers "forked" point-in-time releases of MAME and modified them to work well on the Raspberry Pi.
Its development team started MAME 2003 Plus from a codebase fork based on MAME version 0.78, adding support for newer ROM packages over time (hence, the "Plus" designation).
They built support for the Raspberry Pi and continue to optimize it to run well.
Likewise, the community knows MAME 0.139 as MAME 2010, and a set of developers forked it to make it run well on Raspberry Pi.
However, unlike MAME 2003 Plus, MAME 2010 is not under active development.

Remembering that MAME adds support for new ROM packages over time, we can assume that MAME 2010 supports more ROM packages than MAME 2003 Plus.
However, with MAME 2003 Plus under active development, we can also assume that most ROM packages supported by _both_ MAME 2003 Plus _and_ MAME 2010 perform better in MAME 2003 Plus.
We know a ROM package is supported by both MAME 2003 Plus and MAME 2010 if the ROM package's name is present in both emulators' databases.

However, as the MAME development team releases new versions, they sometimes rename ROM packages or even remove them from the database.
A renamed ROM package can be challenging to identify across different versions of MAME and is beyond the scope of what we cover in this chapter.
However, an ambitious reader may use a "RenameSet" such as [the one hosted by Progretto-Snaps](http://www.progettosnaps.net/renameset/) to identify these renamed ROM packages programmatically.

In general, MAME is distinguished from other emulators by its wide variety of coverage of arcade systems, and its focus on emulation accuracy and historical preservation (e.g., the inclusion of non-playable "mechanical" arcade machines in its database).
If you happen to be reading this chapter and working with a version of MAME (or a fork) newer than version 0.162, then know that MAME also includes support for home console emulation through its merger with another emulator known as MESS.
However, the community generally recommends the use of individual console emulators (i.e., that focus on emulating one console very well) vs. using MAME for home console emulation.

## FBNeo

FBNeo is short for _FinalBurn Neo_.
Its predecessor, _FinalBurn_, was the first emulator to emulate Capcom Play System (CPS) hardware, which made it popular with the emulation community.
The original developer moved on from there to add support for Neo Geo arcade systems.
So the community-at-large knows FinalBurn/FBNeo for its superior emulation of Capcom and Neo Geo systems compared to other emulator options.
Since the original FinalBurn emulator went open source, a new development team forked the codebase and continued development.
The team is still actively adding support for new ROM packages and improving the emulation of existing ones.
The emulator supports both modern, high-performance x86 systems and lower-powered ARM systems like the Raspberry Pi.

FBNeo and MAME overlap in their emulation support for arcade games, meaning that FBNeo, MAME 2003 Plus, and MAME 2010 may emulate a given arcade game.
Luckily for us, the FBNeo team aligned their database to match the ROM package-naming used in MAME's database.
So, if you take a ROM package supported by both FBNeo and MAME and compare the FBNeo database record to that from the most-current MAME release, the ROM package would be listed in both databases, and its name would be identical in each database.

You may be wondering which emulator to use when both FBNeo and MAME support a ROM package.
If the ROM package represents a Capcom or Neo Geo arcade system, FBNeo is almost always the correct choice.
However, if the ROM package does not represent a Capcom or Neo Geo arcade system, the answer is less clear and may depend on real-world testing.
Such testing is beyond the scope of this page, but you may find real-world test results by viewing the [RetroPie MAME wiki page]( https://retropie.org.uk/docs/MAME/) or [FBNeo wiki page](https://retropie.org.uk/docs/FinalBurn-Neo/), and viewing the Compatibility List linked there.

Finally, note that if you happen to be working with the original FinalBurn or a previous fork known as FBAlpha, know that these versions also overlapped in emulation coverage with MAME.
However, the older emulators did not always align their ROM package-naming to that of MAME.
As you will see if you being to perform analysis on ROM packages, misaligned ROM package-naming makes it significantly more challenging to combine two data sets in a join operation.
Therefore, if you are working with Finalburn or FBAlpha instead of FBNeo, you need to use file hashes or otherwise create a ROM package "equivalency lookup table."
