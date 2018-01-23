# Database Generator - VBA
This file is used by `Ecoreal iPM_DB` only.
## Introduction
Database Generator is an Excel file with VBA , which is used to generate Product Database(`Ecoreal iPM_DB`) automatically.

### Product Database
> When All the manual operation is done, this file will create a new sheet named by the 

All kinds of products has their corresponding parameters.
And generally the same series of products has same parameters, but the value of each parameter of the Products belong to a serie is different.  

In the follow picture, showed a series of product called `Config_FS_HZ`.
The header, which is highlighted in orange, is the parameters of `Config_FS_HZ`.

The same `Con_DB`, such as `Config_FS_HZ0001`, means they are corresponding to the same `Reference No.`. (Reference No. is the drawing number of the product) 
![IPM1.png](http://upload-images.jianshu.io/upload_images/9445448-b38079f5a7660f72.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

`Reference No.` is showed in sheets("Ref.").
* ![IPM2.png](http://upload-images.jianshu.io/upload_images/9445448-abb33ab11dec4ace.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

## How to use
* Click `Start` in sheets("HMI")  
![HMI](https://raw.githubusercontent.com/YingjieMA/image/master/VBA/Database%20Generator/HMI.jpg)

![Filter](http://upload-images.jianshu.io/upload_images/9445448-e62a661b1b2f2315.jpg?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

![Generator Tools](https://raw.githubusercontent.com/YingjieMA/image/master/VBA/Database%20Generator/Generator%20Tools.jpg)

##  Remaining problem
Program running slow due to badly designed algorithms.
