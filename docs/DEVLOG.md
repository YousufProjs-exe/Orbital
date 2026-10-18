
# ORBITAL — Devlog

## Goal

Make a tiny satellite survival game inside Microsoft Excel using VBA. No Unity, no Godot, just Excel doing stuff it was never supposed to do lol.
- Just cross the level 6!! 

## Core gameplay
- 3 orbital lanes 
- Move left and right to dodge debris
- Incoming space rocks (they are NOT friendly)
- Energy pickups and scoring
- Difficulty increases as you progress
- Collisions, game over and restart
- Different rockets and flaming debris at higher levels

## Development log

### Session 1 — Getting started
- Created the GitHub Codespace project structure.
- Installed and checked Hackatime tracking.
- Started working out how to make Excel feel like an actual game instead of a boring spreadsheet.

### Session 2 — Making it playable
- Built the basic game board and orbital lanes.
- Added lane switching using the arrow keys and A/D.
- Got the debris moving and added collision detection.
- Started building the scoring, energy and restart systems.

### Session 3 — The game gets harder
- Added level progression and increasing difficulty.
- Kept 3 lanes for the first 5 levels.
- At level 6, the game switches to 2 centered lanes.
- Added a different rocket and flaming debris for the later levels. Things get a bit more serious now.

### Session 4 — Making it look like an actual game
- Added custom pixel-art rockets, space rocks and a space background.
- Worked on the dark space look without making the whole worksheet look empty.
- Added the main menu, mission restart and pause/resume controls.
- Added a menu button so you can return instead of having to mess around with the workbook.
- aaaaa this broken menu - used AI to fix '' 

### Session 5 — Finishing touches
- Fixed the worksheet gridlines and the game area's layout.
- Added the **About Orbital** sheet so people opening the file know what they're looking at.
- Added a Launch Orbital button that opens the game and runs the main menu automatically.
- Fixed the macro startup flow so reviewers don't have to hunt through Excel's macro list just to start playing.
- Cleaned up the final bits and tested the game flow.

## Final result

ORBITAL is finished! A little satellite survival game running entirely in Excel VBA, with custom sprites, levels, lane switching and the whole crash-and-try-again loop.
Was Excel the most normal choice for making a game? Absolutely not.
Was it fun making it work anyway? YES.
Now please don't ask me to build another game in a spreadsheet. (I probably will.)
And if u liked this !? you can see more on my github :) 

## AI Usage ? yes I used but -> 
as ofc VBM was veryyy new for me! 
This project does use AI, but only for coding help when I got stuck, mainly with bugs and figuring out why certain things weren't working as expected. I used it to help debug and get past issues I couldn't solve on my own.

No AI was used to create the assets or decide the creative direction of ORBITAL. The visuals, game design, and overall idea were my own. AI was basically just there when the code decided to ruin my days 😭
