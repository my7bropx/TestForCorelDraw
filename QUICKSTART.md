# Quick Start Guide - CorelDRAW Fill Curve Tool

Get started in 3 minutes!

## Installation (30 seconds)

1. Download `FillCurveTool_Simple.gms`
2. Open CorelDRAW
3. Press `Alt+F11` then `F5`
4. Click `Browse` and select the downloaded file
5. Click `Run` when needed

**That's it!** The macro is now available to use.

## Your First Fill (2 minutes)

### Step 1: Create Objects
```
1. Draw a small circle (e.g., 0.5 inch diameter)
2. Fill it with red color
3. Draw a large rectangle (e.g., 5 x 5 inches)
```

### Step 2: Select Both Objects
```
Click on the small circle while holding Shift
Then click on the large rectangle
(Both should now be selected)
```

### Step 3: Run the Macro
```
Press Alt+F11, then F5
Select "FillCurveSimple" and click Run
```

### Step 4: Configure Settings
```
Pattern: Type "1" (for Grid) and press Enter
Horizontal Spacing: Type "10" and press Enter
Vertical Spacing: Type "10" and press Enter
```

### Step 5: Enjoy!
```
Your rectangle is now filled with red circles!
```

## Common Tasks

### Fill with Logo
```
1. Import/create your logo (small)
2. Create container shape (large)
3. Select both → Run FillCurveSimple
4. Pattern: 2 (Hexagonal), Spacing: 0, 0
```

### Create Watermark Pattern
```
1. Create text "CONFIDENTIAL" (small, rotated)
2. Create rectangle (page size)
3. Select both → Run FillCurveSimple
4. Pattern: 1 (Grid), Spacing: 200, 200
```

### Quick Fill (No Prompts)
```
1. Select fill element + container
2. Run "QuickFill" macro
3. Done! (Uses grid pattern, 0% spacing)
```

## Spacing Explained Simply

Think of spacing as the "breathing room" between elements:

- **0%** = No space (touching)
- **50%** = Half-width space between
- **100%** = Full-width space between
- **-50%** = Overlap by half

## Pattern Types Comparison

| Grid Pattern | Hexagonal Pattern |
|--------------|-------------------|
| ▪ ▪ ▪ ▪      | ▪ ▪ ▪ ▪          |
| ▪ ▪ ▪ ▪      |  ▪ ▪ ▪ ▪         |
| ▪ ▪ ▪ ▪      | ▪ ▪ ▪ ▪          |
| Regular      | Offset/Honeycomb  |

## Keyboard Shortcuts (Optional Setup)

Want to run with one keystroke?

```
1. Tools → Customization → Commands
2. Category: Macros
3. Find "FillCurveSimple"
4. Drag to toolbar
5. Right-click → Properties → Assign shortcut
   (e.g., Ctrl+Shift+F)
```

## Tips for Success

1. **Size matters**: Fill element should be much smaller than container
2. **Start simple**: Try with basic shapes first
3. **Experiment**: Change spacing values to see different effects
4. **Multiple elements**: Select 2+ small shapes for alternating patterns
5. **Undo works**: Don't like it? Just press Ctrl+Z

## Troubleshooting Quick Fixes

| Problem | Quick Fix |
|---------|-----------|
| Nothing happens | Make sure 2+ objects are selected |
| Too many elements | Increase spacing % |
| Too few elements | Decrease spacing % (or use negative) |
| Macro not found | Re-run installation steps |

## Next Steps

Once you're comfortable with the basics:

- Try the **hexagonal pattern** for organic looks
- Use **negative spacing** for overlapping effects
- Experiment with **multiple fill elements** for variety
- Read the full **README.md** for advanced features

## Video Tutorial Concept

*If creating a video tutorial, follow these steps:*

1. **Intro (0:00-0:15)**: Show the end result
2. **Installation (0:15-0:45)**: Quick install demo
3. **Basic Fill (0:45-2:00)**: Circle → Rectangle fill
4. **Spacing Demo (2:00-3:00)**: Show different spacing values
5. **Pattern Comparison (3:00-4:00)**: Grid vs Hex
6. **Advanced (4:00-5:00)**: Multiple elements, rotation
7. **Outro (5:00-5:30)**: Tips and where to get help

---

**Ready to create amazing fills? Start with the basic example above!**

For more details, see **README.md**
