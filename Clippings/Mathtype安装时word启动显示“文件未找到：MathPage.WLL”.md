---
title: "Mathtype安装时word启动显示“文件未找到：MathPage.WLL”"
source: "https://blog.csdn.net/baguette/article/details/136274601"
author:
  - "[[baguette]]"
published: 2024-02-24
created: 2025-09-15
description: "文章浏览阅读6.6k次，点赞32次，收藏59次。本文描述了作者在安装Mathtype过程中遇到的问题，包括不同版本的安装失败、工具栏显示问题和文件缺失。作者详细分享了解决方案，涉及确定电脑位数、文件复制和信任中心设置等步骤，最终成功安装并解决Word兼容问题。"
tags:
  - "clippings"
---
最新推荐文章于 2025-07-30 17:34:48 发布

[塞纳河畔的歌](https://blog.csdn.net/baguette "塞纳河畔的歌") 于 2024-02-24 19:01:10 发布

CC 4.0 BY-SA版权

文章标签： [word](https://so.csdn.net/so/search/s.do?q=word&t=all&o=vip&s=&l=&f=&viparticle=&from_tracking_code=tag_word&from_code=app_blog_art)

版权声明：本文为博主原创文章，遵循 [CC 4.0 BY-SA](http://creativecommons.org/licenses/by-sa/4.0/) 版权协议，转载请附上原文出处链接和本声明。

本文链接： [https://blog.csdn.net/baguette/article/details/136274601](https://blog.csdn.net/baguette/article/details/136274601)

本文描述了作者在安装Mathtype过程中遇到的问题，包括不同版本的安装失败、工具栏显示问题和文件缺失。作者详细分享了解决方案，涉及确定电脑位数、文件复制和信任中心设置等步骤，最终成功安装并解决Word兼容问题。

### 背景

由于老板布置的临时工作，需要 [安装Mathtype](https://so.csdn.net/so/search?q=%E5%AE%89%E8%A3%85Mathtype&spm=1001.2101.3001.7020) ，但尝试了3个不同的版本后（每次都卸载干净了），均未能成功安装，出现的报错3个版本各不相同：

①解压安装过程中失败（这个版本不再尝试）

![](https://i-blog.csdnimg.cn/blog_migrate/cfb29ed6b61e5e56c61c77e5218e6453.png)

②安装后，打开word后的上方 [工具栏](https://so.csdn.net/so/search?q=%E5%B7%A5%E5%85%B7%E6%A0%8F&spm=1001.2101.3001.7020) 不显示“Mathtype”快捷键，并且影响其它软件的正常运行（这个版本尝试3次后选择更换）

![](https://i-blog.csdnimg.cn/blog_migrate/91e2c4a6f6f4e679b02ef893f5b12156.png)

③安装后，word启动显示“文件未找到：MathPage.WLL”（这个是最接近成功的版本，已有同门成功，因此本人进行了进一步的解决问题探索，不到半小时即成功，总结如下）

![](https://i-blog.csdnimg.cn/blog_migrate/69655759e4fa7d8f8cecacc6d9476db4.png)

### 步骤

1\. 确定 **电脑位数** （本人是64位，以下按这个介绍，32位的仅更换相关文件夹即可）

2\. 自行寻找相关版本压缩包并安装（一般搜到我这个帖子的朋友，应该是能够达到这一步的）

3\. 右击桌面上 **MathType的图标** ，点击“ **打开文件所在位置** ”

4\. 找到 **MathPage.WLL** 文件（下附地址）， **复制**

> MathPage.WLL文件位置：... \\ MathType安装地址 \\ MathPage \\ 64 \\ MathPage.wll

![](https://i-blog.csdnimg.cn/blog_migrate/4e03d4edaf30f8255fece6aba669b13d.png)

放到该目录下：

C:\\Program Files\\Microsoft Office \\root\\Office16 **（是的，就是缺这个文件！）**

![](https://i-blog.csdnimg.cn/blog_migrate/487522ad73c6a84ece5649ce0c9d45fb.png)

5\. 找到 **MathType Commands 2016.dotm** 文件（下附地址）， **复制**

> MathType Commands 2016.dotm文件位置：... \\ MathType安装地址 \\ Office Support \\ 64 \\ MathType Commands 2016.dotm

![](https://i-blog.csdnimg.cn/blog_migrate/e5962fe4e495dc0a5235354368fb4810.png)

放到该目录下：

C:\\Program Files\\Microsoft Office\\root\\Office16\\STARTUP **（这里应该存在文件，替换即可）**

![](https://i-blog.csdnimg.cn/blog_migrate/e173072bf019d1a8ca8142c75d4bde55.png)

6\. 现在打开word，工具栏中应该显示“Mathtype”快捷键。点击 **界面左上角的“文件”** ，找到 **“选项”** ，点击 **“信任中心”中的“信任中心设置”**

![](https://i-blog.csdnimg.cn/blog_migrate/e16c7d105a13e9d9ad1db9cb1efd56f7.png)

7\. 添加MathType Commands 2016.dotm所在路径

![](https://i-blog.csdnimg.cn/blog_migrate/4d16f2cf9fc32bf8aa23846d5c73029f.png)

![](https://i-blog.csdnimg.cn/blog_migrate/a53062f94decc224874c4000eb93f432.png)

![](https://i-blog.csdnimg.cn/blog_migrate/c3b4423fa4174d6fed17b5bec190ee51.png)

8\. 重新打开word即可成功加载

![](https://i-blog.csdnimg.cn/blog_migrate/eb3995bff4c9e9da66724f5bdc621df9.png)

关注博主即可阅读全文[*Word* 中 *MathType* 运行 *时* 错误‘53’*:**文件* *未找到**:**MathPage**.**WLL* / 选项灰色无法点击 / Ctrl + V 不能用 / MT EXtra 字体](https://blog.csdn.net/u013669912/article/details/137562003)

[u013669912的博客](https://blog.csdn.net/u013669912)

04-09 1万+[…](https://blog.csdn.net/u013669912/article/details/137562003)[*Mathtype* 无法使用- *word* 中出现error‘48’， *文件* *未找到* ： *MathPage**.**WLL*](https://blog.csdn.net/crrrd/article/details/144828585)

[crrrd的博客](https://blog.csdn.net/crrrd)

12-30 1478[今天打算写毕业论文，因为之前使用的 *mathtype* 是那种30天试用版(因为顺手就有 *安装* 包就没有去官网下载)，然后按照给的方法顺延试用期以后出现了 *文件* *未找到* ： *MathPage**.**WLL* 的问题，导致在 *word* 中无法使用 *Mathtype* 了。最后，总结下遇到这个问题的解决方法，直接搜索找到STARTUP *文件* 夹，然后分别将 *Mathtype* *安装* 路径中的 *MathType* Commands 2016*.*dotm和 *MathPage**.**wll* *文件* 复制到STARTUP *文件* 夹和其上一级 *文件* 夹中(在我的电脑里是Office16)。](https://blog.csdn.net/crrrd/article/details/144828585)[如何解决 *mathpage**.*dll或 *MathType**.*dll *文件* 找不到问题](https://blog.csdn.net/diaozi4616/article/details/101902945)

[diaozi4616的博客](https://blog.csdn.net/diaozi4616)

07-07 496[解决方法（具体图文教程）： 步骤一 要确保路径被office信任。依次打开 *word* -> *文件* ->选项->信任中心->信任中心设置->添加新位置，添加C*:*\\Program Files\\Microsoft Office\\Office14\\STARTUP。 步骤二 在 *MathType* *安装* 目录下找到以下 *文件* (以64为系统为例)： C*:*\\Program Files (x*.**.**.*](https://blog.csdn.net/diaozi4616/article/details/101902945)[运行 *时* 错误‘53’： *文件* *未找到* ： *MathPage**.**WLL* ；please restart *word* to load *mathtype* addin properly](https://blog.csdn.net/qq_51700762/article/details/149782116)

[

最新发布

](https://blog.csdn.net/qq_51700762/article/details/149782116)

[多学会儿的博客](https://blog.csdn.net/qq_51700762)

07-30 534[打开WPS无法粘贴，一直报下面两种错误：运行 *时* 错误‘53’： *文件* *未找到* ： *MathPage**.**WLL* 当前情况是我的 *word* 中没有 *显示* *MathType* ，WPS中 *显示* *MathType* 但实际无法使用，粘贴功能都失效了找了很多教程都是针对 *word* 的修改，最终主包的结果是在，不方便的地方是之后用到 *MathType* 只能 *word* 编辑，好处是主包的WPS终于能正常使用了。](https://blog.csdn.net/qq_51700762/article/details/149782116)[*文件* *未找到* ： *MathPage**.**wll*](https://blog.csdn.net/wwd666666/article/details/140874366)

[wwd666666的博客](https://blog.csdn.net/wwd666666)

08-02 2977[*安装* *mathtype* 后， *文件* *未找到* ： *MathPage**.**WLL* 。](https://blog.csdn.net/wwd666666/article/details/140874366)[*Word* 粘贴 *时* 出现“ *文件* *未找到* ： *MathPage**.**WLL* ”的解决方案](https://blog.csdn.net/weixin_46050242/article/details/142879863)

[weixin\_46050242的博客](https://blog.csdn.net/weixin_46050242)

10-12 8526[一步解决 *word* 中ctrl-v粘贴报错问题](https://blog.csdn.net/weixin_46050242/article/details/142879863)[*word* 中使用 *MathType* *显示* 运行 *时* 错误‘53’,*文件* *未找到* *MathPage**.**wll* 解决方案](https://blog.csdn.net/m0_71604832/article/details/147755274)

[makka的博客](https://blog.csdn.net/m0_71604832)

05-07 3341[网上教程很多，但是结合多个教程完才成功，遂自己记录一下。不报错了，并且也不会 *显示* 按键是灰色的了。](https://blog.csdn.net/m0_71604832/article/details/147755274)[我终于解决 *MathPage**.**wll* *文件* 找不到问题|（最新版 *Word* 上亲测）运行 *时* 错误，53’*:**文件* *未找到**:*athPage*.**WLL*](https://suthel.blog.csdn.net/article/details/136631856)

[群智能算法开发](https://blog.csdn.net/m0_58857684)

03-11 3949[运行 *时* 错误，53’*:**文件* *未找到**:*athPage*.**WLL* 。](https://suthel.blog.csdn.net/article/details/136631856)[在 *Word* 中嵌入 *MathType* *时* 出现“运行 *时* 错误：‘48’（‘53’） *文件* *未找到* ： *MathPage**.**wll* ”的解决方法](https://blog.csdn.net/weixin_53682179/article/details/136371980)

[weixin\_53682179的博客](https://blog.csdn.net/weixin_53682179)

02-29 3024[在 *Word* 中嵌入 *MathType* *时* 出现“运行 *时* 错误：‘48’（‘53’） *文件* *未找到* ： *MathPage**.**wll* ”的解决方法](https://blog.csdn.net/weixin_53682179/article/details/136371980)[*未找到* *MathPage**.**wll* 或 *MathType**.*dll、运行 *时* 错误‘53’： *文件* *未找到* ： *MathPage**.**WLL* *文件* 解决方法](https://blog.csdn.net/weixin_44326071/article/details/128564060)

[weixin\_44326071的博客](https://blog.csdn.net/weixin_44326071)

01-05 9547[*word*](https://blog.csdn.net/weixin_44326071/article/details/128564060)[运行 *时* 错误‘53’ *文件* *未找到* ： *MathPage**.**WLL* ， *安装* *MathType* 后 *Word* 不能复制粘贴问题的解决](https://blog.csdn.net/weixin_53803349/article/details/135310539)

[多多想的博客](https://blog.csdn.net/weixin_53803349)

12-30 7442[我的路径：C*:*\\Program Files\\Microsoft Office\\root\\Office16\\STARTUP\\。我的 *MathPage* *文件* 位置：C*:*\\Program Files (x86)\\ *MathType* \\ *MathPage* \\64。复制放到C*:*\\Program Files\\Microsoft Office\\root\\Office16 *文件* 夹。1*.* 打开 *Word* --> *文件* -->选项-->信任中心-->信任中心设置-->受信任位置，解决宏问题。添加如下受信任位置，](https://blog.csdn.net/weixin_53803349/article/details/135310539)[运行 *时* 错误，53’ *文件* *未找到**:**MathPage**.**WLL*](https://wenku.csdn.net/answer/1mtgc43e0o)

01-07[\### 解决 *MathType* 运行 *时* 错误53 *文件* *未找到* ： *MathPage**.**WLL* 当遇到 *Word* 中的 *MathType* 插件报错“运行 *时* 错误‘53’： *文件* *未找到* ： *MathPage**.**WLL* ”，这通常是由于缺少必要的动态链接库 *文件* 引起的。以下是详细的解决方案：*.**.**.*](https://wenku.csdn.net/answer/1mtgc43e0o)[解决 *MathType* 53号错误： *文件* *未找到* ： *MathPage**.**WLL*](https://blog.csdn.net/CUMT_DJ/article/details/141712830)

[CUMT\_DJ的博客](https://blog.csdn.net/CUMT_DJ)

08-30 1326[应该将 *MathPage**.**wll* 放置到该目录下：C*:*\\Program Files\\Microsoft Office\\root\\Office16。每次装完 *MathType* 后，在 *word* 里面进行粘贴操作 *时* ，总会遇到运行 *时* 错误‘53’： *文件* *未找到* ： *MathPage**.**WLL* 。关于 *启动* *文件* 夹的描述有问题！](https://blog.csdn.net/CUMT_DJ/article/details/141712830)[\[解决方案\]运行 *时* 错误‘53’， *文件* *未找到* ： *MathPage**.**WLL*](https://devpress.csdn.net/v1/article/detail/135579414)

[iii66yy的博客](https://blog.csdn.net/iii66yy)

01-14 2万+[*MathType* Commands 2016*.*dotm 位置：D*:*\\ *mathtype* \\Office Support\\64\\ *MathType* Commands 2016*.*dotm。三、右击 *MathType* 桌面图标，点击“打开 *文件* 所在位置”，然后找到 *MathType* Commands 2016*.*dotm *文件* 所在位置。 *MathPage**.**WLL* 位置：D*:*\\ *mathtype* \\ *MathPage* \\64\\ *MathPage**.**WLL* 。 *mathtype* 使用报错，运行 *时* 错误‘53’， *文件* *未找到* ： *MathPage**.**WLL* 。](https://devpress.csdn.net/v1/article/detail/135579414)[*Word* 加载 *mathtype* 报错找不到 *mathpage**.**wll* *文件*](https://blog.csdn.net/qq_50599660/article/details/136743606)

[qq\_50599660的博客](https://blog.csdn.net/qq_50599660)

03-15 1667[①路径G*:*\\Program Files\\ *MathType* \\Office Support下，根据 *Word* 位数（64）打开名为64的 *文件* 夹，根据 *Word* 版本（2016）选择 *MathType* Commands 2016 *文件* 和 *Word* Cmds *文件* ；②路径G*:*\\Program Files\\ *MathType* \\ *MathPage* 下，根据 *Word* 位数（64）打开名为64的 *文件* 夹，根据 *Word* 版本找到 *MathPage**.**wll* *文件* ；](https://blog.csdn.net/qq_50599660/article/details/136743606)[在 *word* 中嵌入 *MathType* ， *未找到* *MathPage**.**wll* *文件* 的解决方法](https://devpress.csdn.net/v1/article/detail/134615943)

[weixin\_53293346的博客](https://blog.csdn.net/weixin_53293346)

11-25 5958[在 *word* 中嵌入 *mathtype*](https://devpress.csdn.net/v1/article/detail/134615943)[*Word* 粘贴 *时* 出现“ *文件* *未找到* ： *MathPage**.**WLL* ”](https://blog.csdn.net/weixin_43326436/article/details/136159198)

[weixin\_43326436的博客](https://blog.csdn.net/weixin_43326436)

02-18 1735[2*.*需要注意的是， *MathPage**.**WLL* 有32位和64位的，根据自己电脑操作系统复制，如果是64位电脑复制的是64位 *文件* 不好使。那就换一下32位的试一试。找到自己的 *MathType* *文件* 中的 *MathPage**.**WLL* ，再找到Office *安装* 目录，将 *MathType* 中的 *MathPage**.**WLL* 复制进去。一、装完 *MathType* 软件后，用 *Word* *时* ，鼠标粘贴复制可以，用键盘ctrl+v复制不上。1*.*将 *MathType* 软件中的 *MathPage**.**WLL* 复制到Office软件 *安装* 目录下。](https://blog.csdn.net/weixin_43326436/article/details/136159198)[*MathType* 运行 *时* 错误‘53’： *文件* *未找到* ： *MathPage**.**WLL*](https://devpress.csdn.net/v1/article/detail/124925868)

[

热门推荐

](https://devpress.csdn.net/v1/article/detail/124925868)

[m0\_63969219的博客](https://blog.csdn.net/m0_63969219)

05-23 4万+[轻松解决 *MathType* 运行 *时* 错误‘53’： *文件* *未找到* ： *MathPage**.**WLL*](https://devpress.csdn.net/v1/article/detail/124925868)[下载好 *MathPage* 后，打开 *word* ，出现“ *文件* *未找到* ： *MathPage**.**WLL* ”的解决方案](https://blog.csdn.net/2201_75879881/article/details/132740446)

[2201\_75879881的博客](https://blog.csdn.net/2201_75879881)

09-07 2199[可以直接复制这个路径，然后进入到这个路径，后直接粘贴刚刚复制好的*.**wll* *文件* ，如果在 *word* 中使用Math Type *时* 出现这个报错，可以进行以下操作。找 *mathpage**.**wll* *文件* 后，复制这个 *文件* 。然后重新 *启动* wore就完成啦！](https://blog.csdn.net/2201_75879881/article/details/132740446)

登录后您可以享受以下权益：

- 免费复制代码
- 和博主大V互动
- 下载海量资源
- 发动态/写文章/加入社区
×

实付 元

[使用余额支付](https://blog.csdn.net/baguette/article/details/)

点击重新获取

扫码支付

钱包余额 0

抵扣说明：

1.余额是钱包充值的虚拟货币，按照1:1的比例进行支付金额的抵扣。  
2.余额无法直接购买下载，可以购买VIP、付费专栏及课程。

[余额充值](https://i.csdn.net/#/wallet/balance/recharge)

举报

 [![](https://csdnimg.cn/release/blogv2/dist/pc/img/toolbar/Group.png) 点击体验  
DeepSeekR1满血版](https://ai.csdn.net/?utm_source=cknow_pc_blogdetail&spm=1001.2101.3001.10583) ![程序员都在用的中文IT技术交流社区](https://g.csdnimg.cn/side-toolbar/3.6/images/qr_app.png)

程序员都在用的中文IT技术交流社区

![专业的中文 IT 技术社区，与千万技术人共成长](https://g.csdnimg.cn/side-toolbar/3.6/images/qr_wechat.png)

专业的中文 IT 技术社区，与千万技术人共成长

![关注【CSDN】视频号，行业资讯、技术分享精彩不断，直播好礼送不停！](https://g.csdnimg.cn/side-toolbar/3.6/images/qr_video.png)

关注【CSDN】视频号，行业资讯、技术分享精彩不断，直播好礼送不停！

客服 返回顶部

![](https://i-blog.csdnimg.cn/blog_migrate/cfb29ed6b61e5e56c61c77e5218e6453.png) ![](https://i-blog.csdnimg.cn/blog_migrate/91e2c4a6f6f4e679b02ef893f5b12156.png) ![](https://i-blog.csdnimg.cn/blog_migrate/69655759e4fa7d8f8cecacc6d9476db4.png) ![](https://i-blog.csdnimg.cn/blog_migrate/4e03d4edaf30f8255fece6aba669b13d.png) ![](https://i-blog.csdnimg.cn/blog_migrate/487522ad73c6a84ece5649ce0c9d45fb.png) ![](https://i-blog.csdnimg.cn/blog_migrate/e5962fe4e495dc0a5235354368fb4810.png) ![](https://i-blog.csdnimg.cn/blog_migrate/e173072bf019d1a8ca8142c75d4bde55.png) ![](https://i-blog.csdnimg.cn/blog_migrate/e16c7d105a13e9d9ad1db9cb1efd56f7.png) ![](https://i-blog.csdnimg.cn/blog_migrate/4d16f2cf9fc32bf8aa23846d5c73029f.png) ![](https://i-blog.csdnimg.cn/blog_migrate/a53062f94decc224874c4000eb93f432.png) ![](https://i-blog.csdnimg.cn/blog_migrate/c3b4423fa4174d6fed17b5bec190ee51.png) ![](https://i-blog.csdnimg.cn/blog_migrate/eb3995bff4c9e9da66724f5bdc621df9.png)