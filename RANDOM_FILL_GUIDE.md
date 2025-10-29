# CorelDRAW Random Fill Tool - User Guide

## What This Tool Does

Fills any curve (open or closed) with **randomly placed and rotated** elements to create organic, natural-looking patterns with full control over density and spacing.

## Key Features

✅ **Random Placement** - Elements scattered naturally, not in grids
✅ **Random Rotation** - Elements rotate 0-360° for better space filling
✅ **Density Control** - Choose how many elements (50 to 5000+)
✅ **Spacing Control** - Set minimum distance or allow overlap
✅ **Collision Detection** - Optional overlap prevention
✅ **Preserves Properties** - Keeps original colors, size, effects
✅ **Works with Everything** - Open/closed curves, any shape
✅ **CorelDRAW 2018-2026** - Full compatibility

---

## Installation

### Quick Install (30 seconds)

1. Open CorelDRAW
2. Press `Alt + F11` (opens VBA editor)
3. Press `F5` (Play Macro dialog)
4. Click **Browse** → Select `RandomFillTool.gms`
5. Keep this window open for testing

---

## How to Use

### Basic Usage (First Time)

**Step 1: Prepare Your Objects**
```
- Create or select a SMALL element (e.g., small circle, star, logo)
- Create or select a LARGE container (e.g., big circle, rectangle, any curve)
- The container should be much larger than the fill element
```

**Step 2: Select Both**
```
- Click the small element
- Hold SHIFT and click the large container
- Both should be selected (8 handles visible on each)
```

**Step 3: Run the Macro**
```
- In VBA Play Macro window: Select "RandomFillCurve"
- Click "Run"
```

**Step 4: Answer Prompts**
```
1. Fill Density: Enter 500 (for first test)
2. Minimum Spacing: Enter 0 (elements can touch)
3. Allow Rotation: Enter 1 (yes, rotate randomly)
4. Overlap Mode: Enter 1 (allow overlap)
```

**Step 5: Wait**
```
- A "Processing" message appears
- Tool fills the container (may take 5-30 seconds)
- Results appear automatically
```

**Done!** Your container is now filled with randomly placed elements.

---

## Understanding the Settings

### 1. Fill Density (Number of Attempts)

This controls how many times the tool tries to place an element.

| Value | Effect | Best For |
|-------|--------|----------|
| 50-100 | Sparse fill, few elements | Quick tests, minimal look |
| 200-500 | Medium fill, good coverage | Most designs, balanced |
| 1000-2000 | Dense fill, maximum coverage | Packed look, backgrounds |
| 3000+ | Very dense, slow | Special effects |

**Recommendation**: Start with 500, adjust up or down based on results.

### 2. Minimum Spacing (Distance Between Elements)

Controls the gap between elements as a percentage of element size.

| Value | Effect | Visual Result |
|-------|--------|---------------|
| -100% to -50% | Heavy overlap | Elements stack on each other |
| -25% to 0% | Light overlap/touching | Very dense, organic look |
| 0% | Just touching | Maximum coverage, no gaps |
| 10-25% | Small gaps | Natural spacing |
| 50-100% | Large gaps | Airy, spread out |
| 100%+ | Very sparse | Lots of empty space |

**Recommendation**: Use 0% for maximum fill, 10-25% for natural look.

### 3. Allow Random Rotation

**Option 1 (YES)** - Recommended
- Elements rotate 0-360° randomly
- Better space filling
- More organic, natural look
- Helps avoid visible patterns

**Option 2 (NO)**
- Elements keep their original angle
- More uniform appearance
- Faster processing
- Good for directional elements (arrows, text)

**Recommendation**: Use YES (1) for most cases.

### 4. Overlap Mode

**Option 1 (ALLOW)** - Recommended
- Elements can overlap each other
- Much faster (no collision checking)
- Denser fill possible
- Good for: backgrounds, textures, decorative fills

**Option 2 (PREVENT)**
- Checks if elements would collide
- Slower processing (especially with many elements)
- Cleaner, more organized look
- Good for: logos, icons, precise layouts

**Recommendation**: Use ALLOW (1) for dense fills, PREVENT (2) for clean layouts.

---

## Common Use Cases

### Use Case 1: Fill Circle with Scattered Stars
```
Goal: Random star pattern inside a circle

Settings:
- Density: 300
- Spacing: 10%
- Rotation: YES (1)
- Overlap: ALLOW (1)

Result: Natural-looking starfield
```

### Use Case 2: Logo Watermark Background
```
Goal: Subtle logo pattern across entire design

Settings:
- Density: 100
- Spacing: 100%
- Rotation: YES (1)
- Overlap: ALLOW (1)

Result: Sparse, professional watermark
```

### Use Case 3: Dense Confetti Effect
```
Goal: Maximum coverage with colorful shapes

Settings:
- Density: 2000
- Spacing: -25% (overlap)
- Rotation: YES (1)
- Overlap: ALLOW (1)

Result: Dense, overlapping confetti
```

### Use Case 4: Clean Icon Distribution
```
Goal: Icons spread evenly without overlap

Settings:
- Density: 500
- Spacing: 25%
- Rotation: NO (2)
- Overlap: PREVENT (2)

Result: Organized icon pattern
```

### Use Case 5: Organic Leaf Pattern
```
Goal: Natural leaves filling a branch shape

Settings:
- Density: 800
- Spacing: 0%
- Rotation: YES (1)
- Overlap: ALLOW (1)

Result: Natural, organic foliage
```

---

## Tips for Best Results

### 1. Element Size
- Fill elements should be 10-20x smaller than container
- Too large = few elements placed
- Too small = may take very long

### 2. Density vs Speed
- Higher density = more elements but slower
- Start low (200-300) and increase if needed
- Over 2000 may take 1-2 minutes

### 3. Multiple Elements
- Select multiple different shapes to fill with
- Tool will randomly cycle through them
- Creates more varied, interesting patterns

### 4. Testing
- Use **QuickRandomFill** for instant tests (no prompts)
- Preset: 1000 density, rotation ON, overlap ON
- Perfect for quick previews

### 5. Undo
- Entire fill is one undo step
- Don't like it? Press **Ctrl+Z**
- Try different settings

### 6. Performance
- Allow overlap = Much faster
- Prevent overlap = Slower but cleaner
- For 1000+ elements, be patient

---

## Quick Reference

### Macro Names

| Macro | Description |
|-------|-------------|
| `RandomFillCurve` | Main tool with full options |
| `QuickRandomFill` | Instant fill (no prompts, preset: dense with rotation) |

### Recommended Presets

**Sparse Natural Fill**
- Density: 300
- Spacing: 20%
- Rotation: YES
- Overlap: ALLOW

**Dense Packed Fill**
- Density: 1500
- Spacing: -10%
- Rotation: YES
- Overlap: ALLOW

**Clean Organized Fill**
- Density: 500
- Spacing: 25%
- Rotation: NO
- Overlap: PREVENT

**Maximum Coverage**
- Density: 2000
- Spacing: 0%
- Rotation: YES
- Overlap: ALLOW

---

## Troubleshooting

### Problem: Very few elements placed

**Solutions:**
- Increase density (try 1000-2000)
- Decrease spacing (try 0% or negative)
- Enable overlap mode (option 1)
- Make sure container is much larger than elements

### Problem: Tool is too slow

**Solutions:**
- Reduce density (try 300-500)
- Enable overlap mode (faster)
- Use smaller fill elements
- Close other programs for more memory

### Problem: Elements outside container

**Solutions:**
- This shouldn't happen - tool checks boundaries
- Make sure container is a valid curve
- Try converting container to curves manually first

### Problem: All elements same rotation

**Solutions:**
- Make sure you selected "1" for rotation question
- Check that elements aren't grouped
- Try running QuickRandomFill instead

### Problem: Macro doesn't run

**Solutions:**
- Check CorelDRAW version (needs 2018+)
- Enable macros in Tools > Options > VBA
- Try running from VBA editor (Alt+F11, F5)
- Make sure at least 2 objects selected

---

## Technical Notes

### How It Works

1. **Selection**: Identifies largest object as container
2. **Random Generation**: Creates random X, Y positions
3. **Boundary Check**: Verifies point is inside container curve
4. **Collision Check**: (if enabled) Checks distance to other elements
5. **Placement**: Duplicates element, positions, rotates
6. **Repeat**: Continues for specified number of attempts

### Performance

- **Fast Mode** (overlap allowed): ~1000 elements in 10-20 seconds
- **Slow Mode** (collision detection): ~500 elements in 20-40 seconds
- Depends on computer speed and element complexity

### Limitations

- Maximum practical: ~5000 elements (may take several minutes)
- Very complex curves may slow down boundary checking
- Collision detection slows down with many elements

---

## Examples Gallery

### Example 1: Starry Night
```
Element: Small yellow star
Container: Large circle
Settings: Density 500, Spacing 15%, Rotation YES, Overlap ALLOW
Result: Beautiful starfield effect
```

### Example 2: Confetti Celebration
```
Elements: 3 different colored rectangles (small)
Container: Large rectangle
Settings: Density 1500, Spacing -20%, Rotation YES, Overlap ALLOW
Result: Dense overlapping confetti
```

### Example 3: Logo Watermark
```
Element: Company logo (small, semi-transparent)
Container: Page rectangle
Settings: Density 200, Spacing 150%, Rotation NO, Overlap ALLOW
Result: Subtle watermark pattern
```

### Example 4: Organic Texture
```
Element: Irregular blob shape
Container: Freeform curve
Settings: Density 1000, Spacing 0%, Rotation YES, Overlap ALLOW
Result: Natural organic texture
```

---

## Keyboard Shortcut Setup (Optional)

To run with one keystroke:

1. **Tools** → **Customization** → **Commands**
2. Category: **Macros**
3. Find **RandomFillCurve**
4. **Drag** to toolbar
5. Right-click icon → **Properties**
6. Assign shortcut (e.g., **Ctrl+Shift+R**)

---

## Need Help?

- Start with **QuickRandomFill** for instant results
- Use low density (200-300) for testing
- Read "Common Use Cases" section for inspiration
- Experiment with different settings - undo always works!

---

**Enjoy creating beautiful random patterns!**
