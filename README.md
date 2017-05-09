InsertImage++
======

[InsertImage++](https://github.com/kennywei815/InsertImagePlus) is a Microsoft Office add-in which makes Word, Excel, and PowerPoint support most image file formats (via ImageMagick convert).

- Support all MS Office products: Word, Excel, and PowerPoint.
- Support MS Office 2003 to Office 2016.
- Mac OS Support is also in massive development.

# Prerequisite
You may need to install these packages first:
- [Ghostscript](https://www.ghostscript.com/download/gsdnld.html) for PDF, PS file processing

# Quick Start

### Install
1. `git clone https://github.com/kennywei815/InsertImagePlus.git`
2. `cd InsertImagePlus/`
3. `.\setup.bat`

### Usage

Select "Insert/Change Image" from the "Add-Ins" tab of the ribbon, and you will get a file dialog where you can load your image:

![alt text](https://github.com/kennywei815/InsertImagePlus/blob/master/Office_Add-Ins.png)

![alt text](https://github.com/kennywei815/InsertImagePlus/blob/master/FileDialog.png)

Choose your image, and click on "Open". InsertImage++ will convert your image into PNG format, and insert it into Office.

![alt text](https://github.com/kennywei815/InsertImagePlus/blob/master/Result.png)

If you need to change the image, just select the image, then click on "Insert/Change Image" again, and the file dialog will re-appear so you can load another image.

You can also treat the loaded image as an ordinary PowerPoint image. For example, it can be grouped, animated, rotated, moved, and resized. Further editing of the loaded image will preserve all these changes.

# License
Copyright (c) 2017 Cheng-Kuan Wei Licensed under the Apache License.

This software contains code derived from Jonathan Le Roux and Zvika Ben-Haim's IguanaTeX project which is originally released under the Creative Commons Attribution 3.0 License and combined into InsertImage++ as a whole under Apache License 2.0 .