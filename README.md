# Database Generator - VBA
This file is used by `Ecoreal iPM_DB` only.
## Introduction
Database Generator is an Excel file with VBA , which is used to generate Product Database(`Ecoreal iPM_DB`) automatically.

### Product Database

All kinds of products has their corresponding parameters.
And generally the same series of products has same parameters, but the value of each parameter of the Products belong to one serie is different.  

In the follow picture, showed a series of product called `Config_FS_HZ`.
The header, which is highlighted in orange, is the parameters of `Config_FS_HZ`.

The same `Con_DB`, such as `Config_FS_HZ0001`, means they are corresponding to the same `Reference No.`. (Reference No. is the drawing number of the product) 
![IPM1.png](http://upload-images.jianshu.io/upload_images/9445448-b38079f5a7660f72.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

`Reference No.` is showed in sheets("Ref.").

![IPM2.png](http://upload-images.jianshu.io/upload_images/9445448-abb33ab11dec4ace.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

## How to use

*  Click `Start` in sheets("HMI")  to start the `Generator Tools`.
![HMI](https://raw.githubusercontent.com/YingjieMA/image/master/VBA/Database%20Generator/HMI.jpg)


* Parameters and their value is listed in the left listview. Actually I list the all parameters you may use in `Ecoreal iPM_DB`.
![Generator Tools](https://raw.githubusercontent.com/YingjieMA/image/master/VBA/Database%20Generator/Generator%20Tools.jpg)

*  Double click the item, if you want to change the value. You can choose the value you want or not in the form `Filter`. Click `OK` to confirm.
![Filter](http://upload-images.jianshu.io/upload_images/9445448-e62a661b1b2f2315.jpg?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)


*  Choose one item then click `Clear`(the bottom on the left_top of the  `Generator Tools`), so that you can delete the parameter you do not need.
![process 3.png](http://upload-images.jianshu.io/upload_images/9445448-e554b4ab093ac582.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)


* Double click items in listview `Ref Select` to choose the corresponding `Reference No.` . Here is already set a library system, so that you can search the  `Reference No.` easily. And you also can change the quantity etc.
![process 4.png](http://upload-images.jianshu.io/upload_images/9445448-f68214f71f8359fd.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)



*  When all the options are selected, fill the `Con_DB` in the left_bottom of the `Generator Tools`. Then click the button `Generator`. A new sheet named by what you filled in `Configuration` is created, and the database is already generated.
![process 2.png](http://upload-images.jianshu.io/upload_images/9445448-bc5bbbcb796642a0.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

##  Remaining problem
Program running slow due to badly designed algorithms.
