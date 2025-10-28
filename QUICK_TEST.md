# Quick Test - 2 Minutes

## Test the Random Fill Tool Right Now

### Step 1: Open CorelDRAW (any version 2018-2026)

### Step 2: Create Test Objects
```
1. Draw a small circle (about 0.5 inch / 1 cm diameter)
2. Fill it with any color (e.g., red)
3. Draw a large rectangle (about 5x5 inches / 12x12 cm)
4. Keep rectangle unfilled or use light color
```

### Step 3: Select Both Objects
```
- Click the small circle
- Hold SHIFT
- Click the large rectangle
- Both should show selection handles
```

### Step 4: Run the Macro
```
1. Press Alt+F11 (opens VBA editor)
2. Press F5 (Play Macro dialog)
3. Click "Browse"
4. Select "RandomFillTool.gms"
5. Select "RandomFillCurve" in the list
6. Click "Run"
```

### Step 5: Enter Settings (First Test)
```
Dialog 1 - Density:
Type: 500
Press Enter

Dialog 2 - Spacing:
Type: 0
Press Enter

Dialog 3 - Rotation:
Type: 1
Press Enter

Dialog 4 - Overlap:
Type: 1
Press Enter
```

### Step 6: Wait for Result
```
- "Processing" message appears
- Wait 10-20 seconds
- Result message shows number of elements placed
- Your rectangle is now filled with randomly scattered circles!
```

---

## If It Worked - Try These Next:

### Test 2: More Dense
```
Same as above, but:
- Density: 1000
- Spacing: -25 (overlap)
Result: Very dense, overlapping circles
```

### Test 3: Sparse with Gaps
```
Same as above, but:
- Density: 200
- Spacing: 50
Result: Few circles with large gaps
```

### Test 4: Quick Fill (No Prompts)
```
1. Select objects
2. Run "QuickRandomFill" macro
3. Instant result with preset settings!
```

---

## Settings Cheat Sheet

| Setting | What to Enter | Result |
|---------|---------------|--------|
| Density | 500 | Medium fill |
| Density | 1500 | Very dense |
| Spacing | 0 | Touching |
| Spacing | -50 | Overlap 50% |
| Spacing | 50 | Medium gaps |
| Rotation | 1 | Rotate randomly |
| Rotation | 2 | Keep angle |
| Overlap | 1 | Allow overlap (fast) |
| Overlap | 2 | Prevent collision (slow) |

---

## Troubleshooting First Test

**Nothing happens:**
- Make sure 2 objects selected
- Check VBA is enabled (Tools > Options > VBA)

**Very few circles placed:**
- Increase density to 1000
- Make rectangle much larger

**Too slow:**
- Reduce density to 300
- Make sure Overlap is 1 (allow)

**Undo:**
- Press Ctrl+Z to undo entire fill
- Try different settings

---

**That's it! See RANDOM_FILL_GUIDE.md for complete documentation.**
