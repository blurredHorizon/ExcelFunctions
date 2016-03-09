##VLOOKUP2

###Rationai:
Sometimes the provided VLOOKUP function will not find search items because it coerces the entered and stored search key into different types.  This function ensures that the comparison will take place between objects of the same type: Strings.

###Use:

[] | A | B
---| --- | ---
1> | 4 | a
2> | 5 | b
3> | 6 | c

```
'prints "a" in cell
=VLOOKUP2(A1, "Sheet1", "A1:A3", 1)
```

