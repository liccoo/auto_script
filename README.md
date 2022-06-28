
>   首先解压文件，解压后<font color=red>务必保持`/image/`、`mian.py`、`command.xls`文件的相对路径关系，否者该程序将无法运行</font>。`command.xls`文件记录的是流程自动化软件（RPA）的运行步骤，`/image/`文件夹下面存放的就是鼠标点击的图标或区域的截图，软件或代码的运行机制为从文件`command.xls`读取软件的运行步骤，然后按照步骤依次执行。

## 1.环境配置(exe版本直接跳过)  --- 请自行打包

*   源码版本必须依赖python3，我用的是python3.7，更高版本未经测试，但应该也可以

*   安装依赖包 --- 下载包总是失败慢的话，用镜像源安装

    ```
    pip install xlrd
    pip install pyautogui
    pip install opencv-python
    pip install pillow
    pip install pyperclip
    ```

## 2.操作步骤

1.   在`command.xls`的`Sheet1`表格逐步地进行配置，从第二行开始，配置好了以后保存文件。各部分的详细作用及用法如下表。

|  列号  | 说明                                                         |
| :----: | :----------------------------------------------------------- |
| 第一列 | （<font color=red>必填项</font>）代表操作指令类型，只能填 1~6 的自然数，（1 单击  2 双击  3 右键  4输入  5 等待  6滚轮） |
| 第二列 | （<font color=red>必填项</font>）对应操作指令的操作内容，1.若是鼠标点击操作，则内容是图片名称；2.若是输入超作，则内容是普通的输入文本或键盘功能键代码；3.若是等待操作，则内容是等待时长，单位是秒(s)；4.若是滚轮操作，则内容是滚动格值（建议取所需滚动像素的两倍），正数表示向上滚，负数表示向下滚。 |
| 第三列 | （<font color=red>选填项，默认为1</font>）重复次数，只能取-1或大于0的自然数 |
| 第四列 | （<font color=red>选填项，不会影响程序运行</font>）操作说明  |

2.   把每一步要操作的<font color=red>鼠标点击区域截图</font>保存至`/image/`文件夹下，一定得是`.png` 格式（<font color=red>别用中文命名</font>，建议以序号命令）。

>   注意：截图的原则必须满足<font color=red>唯一性</font>，若为鼠标点击任务，默认鼠标光标定位坐标为所截图像的中心坐标，所以如果同一屏幕上有多个相同图标，会默认找到最左上的一个。因此怎么截图，截多大的区域，在截图的时候就得考虑好，如输入框只截中间空白部分应该是不行的（可能没法定位）。

3.   `exe`版双击`main.exe`即可启动，源码版运行程序即可（务必保持`image`、`mian.py`、`command.xls`文件的相对路径关系，否者该程序将无法运行）。

## 3.Tips

1.   程序开始运行后把鼠标光标移至屏幕的最左上角可中途终止程序
2.   程序在开始运行和结束运行时都会有弹窗提示
3.   若不会Python或嫌麻烦，推荐使用exe版本，MacOS和Linux需自行用源码打包或采用源码模式。
4.   程序开始后可移动光标暂停程序，确认后继续运行，注意<font color=red>大部分软件在扑获光标前后界面会有变化，由此可能会造成截图失效（建议继续前手动让暂停前的应用获得光标）</font>。
5.   如果需要等待前一步任务执行完毕再执行下一步任务，可能造成任务无法往下继续进行（程序会一直匹配图片直到匹配到为止，与报错冲突，暂时没有很好的解决方案）
6.   `Sheet2`为模板

## 4.功能键参考代码

```
# 输入''或""里d
['\t', '\n', '\r', ' ', '!', '"', '#', '$', '%', '&', "'", '(', ')', '*', '+', ',', '-', '.', '/', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', ':', ';', '<', '=', '>', '?', '@', '[', '\\', ']', '^', '_', '`', 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '{', '|', '}', '~', 'accept', 'add', 'alt', 'altleft', 'altright', 'apps', 'backspace', 'browserback', 'browserfavorites', 'browserforward', 'browserhome', 'browserrefresh', 'browsersearch', 'browserstop', 'capslock', 'clear', 'convert', 'ctrl', 'ctrlleft', 'ctrlright', 'decimal', 'del', 'delete', 'divide', 'down', 'end', 'enter', 'esc', 'escape', 'execute', 'f1', 'f10', 'f11', 'f12', 'f13', 'f14', 'f15', 'f16', 'f17', 'f18', 'f19', 'f2', 'f20', 'f21', 'f22', 'f23', 'f24', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9', 'final', 'fn', 'hanguel', 'hangul', 'hanja', 'help', 'home', 'insert', 'junja', 'kana', 'kanji', 'launchapp1', 'launchapp2', 'launchmail', 'launchmediaselect', 'left', 'modechange', 'multiply', 'nexttrack', 'nonconvert', 'num0', 'num1', 'num2', 'num3', 'num4', 'num5', 'num6', 'num7', 'num8', 'num9', 'numlock', 'pagedown', 'pageup', 'pause', 'pgdn', 'pgup', 'playpause', 'prevtrack', 'print', 'printscreen', 'prntscrn', 'prtsc', 'prtscr', 'return', 'right', 'scrolllock', 'select', 'separator', 'shift', 'shiftleft', 'shiftright', 'sleep', 'stop', 'subtract', 'tab', 'up', 'volumedown', 'volumemute', 'volumeup', 'win', 'winleft', 'winright', 'yen', 'command', 'option', 'optionleft', 'optionright']
```

## 5.pyinstaller 打包成 exe
* 安装 pyinstaller --- pip install pyinstaller or pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple/
* cmd 切换到所需打包的脚本文件目录
* 执行打包命令，例如：pyinstaller -F -w -i python.ico main.py，中间的这些参数根据自己的需求进行选择，常用 -F,-w,-i 
  * -F：代表每次打包会进行覆盖操作，生成结果是一个 exe 文件，所有的第三方依赖、资源和代码均被打包进该 exe 内 
  * -w：不显示命令行窗口，根据自己的需求进行选择，（仅对 Windows 有效） 
  * -c：显示命令行窗口（默认），（仅对 Windows 有效） 
  * -i：指定图标，就是打包带有自己设置的 ico 图标 
  * -D：生成结果是一个目录，各种第三方依赖、资源和 exe 同时存储在该目录（默认） 
  * -d：执行生成的 exe 时，会输出一些log，有助于查错 
  * -v：显示版本号
  
* 打包好的程序会在生成的 dist 目录下

我们写的python脚本是不能脱离python解释器单独运行的，所以在打包的时候，至少会将python解释器和脚本一起打包，同样，为了打包的exe能正常运行，会把我们所有安装的第三方包一并打包到exe。

例如，我们的项目只使用的一个requests包，但是可能我们还安装了其他n个包，但是他不管，因为包和包只有依赖关系的。比如我们只装了一个requests包，但是requests包会顺带装了一些其他依赖的小包，所以为了安全，只能将所有第三方包+python解释器一起打包。

> 参考文章链接：https://blog.csdn.net/weixin_46348230/article/details/120684653 <br>
> 参考文章链接：https://zhuanlan.zhihu.com/p/162237978