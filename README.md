# CIEDE2000-Color-Difference

This repository provides an implementation of the [CIEDE2000 Color-Difference formula](http://www2.ece.rochester.edu/~gsharma/ciede2000/) (de00) for Mathematica, C# and Excel.
Its purpose is to calculate colors in the ab-place of the Lab color space that have a defined color-disance when L is fixed.

![Screenshot](screenshot.png)

The CIEDE2000 formula has several improvements, e.g. it deals with the problematic blue region. The goal is to get colors around a reference color that have a specific color-difference. In Mathematica, we can use `ContourPlot` to easily visualize this. To calculate it, we use simple scheme where we use a reference color and a certain direction, and we calculate at which distance we approach the sought color-difference.
So we try to find the root

```
0 = de00(r) - distance
```

The repository contains 

1. A prototype implementation of de00 to verify the correctness of the formula and to do some initial testing
2. A C# translation of the code and an algorithm to find above root along a defined direction
3. Bindings for Excel using [Excel-DNA](https://github.com/Excel-DNA) that allows to call the functionality directly as formula inside Excel.
