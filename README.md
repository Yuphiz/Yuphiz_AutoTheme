# Yuphiz_AutoTheme

0、隐私问题
本脚本默认用百度api定时经纬度，可以自己手动设置，设置见下面的简单设置

1、工具描述
本脚本套可以提高 Windows10 深浅色主题的切换体验


2、简单使用方法：

![image](https://github.com/Yuphiz/Yuphiz_AutoTheme/blob/main/demo%E9%A2%84%E8%A7%88%E5%9B%BE/%E8%87%AA%E5%8A%A8%E4%B8%BB%E9%A2%98%E4%BD%BF%E7%94%A8%E5%9B%BE%E7%AE%80%E5%8D%95%E5%BF%AB%E9%80%9F%E4%BD%BF%E7%94%A8%E5%9B%BE.gif)

更多设置和功能需要下载其他扩展配合使用，扩展正在逐步上传



3、使用预览图，以下演示都需要单独下载扩展和选择后台版才有相同效果

![image](https://github.com/Yuphiz/Yuphiz_AutoTheme/blob/main/demo%E9%A2%84%E8%A7%88%E5%9B%BE/%E4%BD%BF%E7%94%A8demo1_%E5%A3%81%E7%BA%B8%E6%89%A9%E5%B1%95.gif)

3.1 使用demo1_壁纸和扩展


![image](https://github.com/Yuphiz/Yuphiz_AutoTheme/blob/main/demo%E9%A2%84%E8%A7%88%E5%9B%BE/%E4%BD%BF%E7%94%A8demo2_UWP%E5%90%AF%E5%8A%A8%E9%A1%B5%E9%A2%9C%E8%89%B2.gif)

3.2 使用demo2_UWP启动页颜色


![image](https://github.com/Yuphiz/Yuphiz_AutoTheme/blob/main/demo%E9%A2%84%E8%A7%88%E5%9B%BE/%E4%BD%BF%E7%94%A8demo3_%E5%9B%BE%E6%A0%87%E6%BC%94%E7%A4%BA.gif)

3.3 使用demo3_图标演示


![image](https://github.com/Yuphiz/Yuphiz_AutoTheme/blob/main/demo%E9%A2%84%E8%A7%88%E5%9B%BE/%E4%BD%BF%E7%94%A8demo4_%E6%98%BE%E7%A4%BA%E5%99%A8%E4%BA%AE%E5%BA%A6%E5%AF%B9%E6%AF%94%E5%BA%A6%E6%A6%82%E5%BF%B5%E9%A2%84%E8%A7%88.gif)

3.4 使用demo4_显示器亮度对比度概念

此为演示图，实际没有这个界面，后台自动调节
注：无论内置还是外接显示都需要支持ddc/ci，内置显示器不支持调对比度，外接显示器支持ddc/ci协议的可以调对比度和亮度


![image](https://github.com/Yuphiz/Yuphiz_AutoTheme/blob/main/demo%E9%A2%84%E8%A7%88%E5%9B%BE/%E4%BD%BF%E7%94%A8demo5_%E7%AC%AC%E4%B8%89%E6%96%B9%E8%BD%AF%E4%BB%B6.gif)

3.5 使用demo5_第三方软件演示（兼容性视本身软件支持）




4、简单设置

4.1、如果没网或者网络不佳导致不能定位（不能定位就是东经0，纬度0），可以使用下面方法手动定位

![image](https://github.com/Yuphiz/Yuphiz_AutoTheme/blob/main/demo%E9%A2%84%E8%A7%88%E5%9B%BE/q%26a%201%20%E5%AE%9A%E4%BD%8D%E9%97%AE%E9%A2%98.png)

打开脚本所在位置，找到下图红线框住的文件

![image](https://github.com/Yuphiz/Yuphiz_AutoTheme/blob/main/demo%E9%A2%84%E8%A7%88%E5%9B%BE/q%26a%201%20%E5%AE%9A%E4%BD%8D%E9%97%AE%E9%A2%98%E8%A7%A3%E7%AD%941.jpg)

用编辑器打开，没有就直接用记事本，这里以记事本为例

![image](https://github.com/Yuphiz/Yuphiz_AutoTheme/blob/main/demo%E9%A2%84%E8%A7%88%E5%9B%BE/q%26a%201%20%E5%AE%9A%E4%BD%8D%E9%97%AE%E9%A2%98%E8%A7%A3%E7%AD%942.png)

找到【网络自动定位】，把”开”改成”关”，然后网上找一下自己所在位置的经纬度，填进去
比如东经123，南纬12
就填
,     "手动定位经度":123
,     "手动定位纬度":-12
保存文档（注意编码是ansi（中文系统下），或者gb2312），再重新按使用方法使用



4.2、壁纸所在位置是 【脚本所在位置】\扩展\壁纸高级版_Yuphiz\壁纸包

支持jpeg、png、bmp、jpg图片格式，浅色主题壁纸以_1结尾，深色主题壁纸以_2结尾。

特别注意的是，每一组图片文件夹只能有1张_1，1张_2图片，多了就不能读取。所以每一组图片要放到单独的文件夹内。更多设置需要在脚本单独页查看

![image](https://github.com/Yuphiz/Yuphiz_AutoTheme/blob/main/demo%E9%A2%84%E8%A7%88%E5%9B%BE/q%26a%202%20%E5%A3%81%E7%BA%B8.png)




5、个人开发不易，如果觉得解决了你的问题，请捐赠支持开发者。你的支持将会让工具越来越好。如果多人支持，后期会考虑出个脚本ui版

![image](https://github.com/Yuphiz/Yuphiz_AutoTheme/blob/main/demo%E9%A2%84%E8%A7%88%E5%9B%BE/Yuphiz_Pay.jpg)
