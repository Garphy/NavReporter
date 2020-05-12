# NavReporter
A tool automatically scans emails, downloads Excel files,  parse the spreadsheet, collect the data and send them to related persons.
1. 自动扫描邮件，下载指定附件。
2. 解析Excel电子表格，提取感兴趣信息。
3. 自动填写Excel电子表格，生成报告。
4. 自动发送报告到指定电子邮件账户。

最新修改：

1. 下载的代码可以直接运行。使用测试邮箱 nav_report@163.com，请勿修改邮箱设置和密码。
2. 163.com 邮箱默认不开启pop3功能，如果自己编写程序，记得设置开启pop3收信功能。
3. 代码中Calendar.py中第32行被hardcode成 2019.9.19，在实际使用中需要修改。
4. 目前程序做成了一个实例，程序会登录nav_report@163.com，下载两个邮件附件，里面是模拟的数据表，然后读取数据，填写本地的报告表格，再把报告发送回nav_report@163.com。
5. 程序运行两种方式，假设下载代码后，放置于 NavReporter目录：
  1). 运行方法一： python_path/python.exe NavReporter ，这种运行方法会执行 NavReporter/__main__.py中的入口
  2). 运行方法二： python_path/python.exe NavReporter/NavReporter.py，这种方法会执行NavRepoerter/NavReporter.py中的__main__入口。

# zhihu
关键词：Python, 邮件扫描， yaml配置文件，邮件附件编解码，Excel自动处理 



**一 、前言**

​       作为一只二级市场资管狗，平时少不了向各相关方汇报产品数据，不论领导、风控部门同事、市场人员或者客户，经常会要求我们按照一定频率披露相关数据（工作年头一长，感觉需要汇报的人越来越多![img](https://res.wx.qq.com/mpres/htmledition/images/icon/common/emotion_panel/emoji_wx/2_05.png)）。资管产品最准确的数据来源就是每个交易日收盘后，经过估值清算生成的估值表。但是这些估值服务供应商不同，表格格式差别较大，内容也比较繁杂，不论从可读性和信息保密方面考虑，都不应该随意将估值表提供给对方。比较合适的做法是将对方感兴趣的且允许披露的内容提取出来，放到一个汇总报告中。

​       靠手工完成这类简单重复工作比较耗时，且容易让人感觉枯燥导致出错，这里我用Python实现了一个工具，能够自动扫描邮件，下载Excel电子表格附件，抽取所需信息并汇总，再将报告自动发送到指定邮件帐户。看似简单的功能，里面还需要不少技巧，完工后顺便写篇帖子总结一下经验。



**二、系统设计**

​      图1. 所示是系统设计。因为是净值（Net Asset Value）报告程序，整个程序的入口类叫做NavReporter。顶层模块下面实现了6个功能模块，每个功能模块都以单例（Singleton）模式实现。分别完成配置信息解析（Config），日期文件解析和简单日期函数（Calendar），邮箱扫描定位目标邮件并下载附件（EmailScanner），估值表电子表格解析和信息提取（NavParser），汇总报告表格自动填写（ReportFiller）以及报告发送（ReportSender）等几块功能。分成多模块实现目标是每个类尽量争取做到多带带可用，这样邮件附件下载后，可以根据不同需求解析不同信息，填写不同格式报告，发送给不同接收人。

![img](http://img1.qdcypf.com/mmbiz_jpg/ANUAFkOeAz0IaeZHIbq1pqwWlRCDUVcBy1KSZceKbQbibF0JGgYup0tr0KITvXuXgJ8394sHjz6WTD5PIyvv1vQ/640?wx_fmt=jpeg)

​                    图1. 系统结构图



**三、功能实现**

   \1. Config模块

​       这个模块实现比较简单，主要使用了yaml库解析yaml格式的配置文件。我原本打算使用json格式文件作为配置文件，但是发现手动编辑的json文件总是出现莫名的解析错误。看到一些经验分享文章也不推荐使用json作为配置文件。json本身设计目的是作为一种数据交换格式，对于注释信息支持较差，自身语法格式规范也很严格（大概是我手动编辑引发解析错误的原因），不太适合作为配置文件使用。

​      这里推荐使用yaml库，文件格式结构与json接近，层次关系省略了json大量使用的花括号，代之以缩进方式。对于yaml格式配置文件，yaml.load()一个方法就能把文件读入，并解析成字典格式，非常方便。

​       yaml库解析出来的内容，会放到一个字典结构中，因为配置项存在嵌套关系，返回的字典也会是嵌套结构，访问一个深层次配置项会有 config_info['TopLev']['MidLev']['BotLev']这种情况。针对此我引入属性字典结构，把这种嵌套字典都映射成不同层次的属性，提高了可读性，至少也可以省略多个中括号和引号。映射成属性字典后，访问一个层次嵌套属性会变成如下形式：config_info.TopLev.MidLev.BotLev。实现这种功能只需要一个深度优先递归遍历字典结构并更新Python的类成员self.__dict__操作。图2是这部分的关键代码。



![img](http://img1.qdcypf.com/mmbiz_png/ANUAFkOeAz0IaeZHIbq1pqwWlRCDUVcBmA2ufMcAvdZg3akJWQIQAYicO2mYkPqLibiaZQibOD2Su7icruCsgGxxhgg/640?wx_fmt=png)

图2. 递归更新self.__dict__成员实现属性字典



   \2. Calendar模块

​       日期模块功能也很简单。日期模块功能依赖于初始化时候读入事先设置好的日历文件，包含了交易日信息，通常只有在交易日才会有新的估值信息。日期模块提供了交易日判断功能，判断传入的日期是否是交易日，如果不是，程序可以直接退出，无需做任何工作。另外，这个模块还提供了方法获得指定的交易日。假设当前是T日，多数情况下，T日只能获得T-1日或者T-2日的估值信息，所以模块提供了get_prev_trading_day()方法获得前一个交易日信息。



​    \3. EmailScanner模块

​       邮件扫描先要登录邮箱，通过配置文件提供的POP3服务器地址、用户名和密码，调用Python的poplib库就可以登录邮件服务器，获得邮箱内邮件信息。图3所示登录代码片断。



![img](http://img1.qdcypf.com/mmbiz_jpg/ANUAFkOeAz0IaeZHIbq1pqwWlRCDUVcBXRpCrHk5KZ3EkpHcSa5icaaEXXj2KGP3tRrh4U9ibu3iafIblWP33zdpQ/640?wx_fmt=jpeg)

图3. 登录POP3服务器



​      登录成功后，按照由新及旧的顺序，逐个扫描邮件。先解析邮件头信息，然后查看发件人，如果发件人不是估值表提供商，则循环到下一个邮件。匹配发件人成功的邮件，还要检查是否存在附件，如果存在附件，则检查附件名是否是所需要文件，只有以上规则都匹配，才将附件下载，否则就跳到下一个邮件。提取附件时候，邮件解码是个容易出错问题，具体内容在下一节避坑指南中介绍。图4是邮件扫描流程图。



![img](http://img1.qdcypf.com/mmbiz_jpg/ANUAFkOeAz0IaeZHIbq1pqwWlRCDUVcBtJNmOQhcEGCq4To03ID1O3aFNWtiaePiceOI3xLlsvchwEBnZA8pia5icw/640?wx_fmt=jpeg)

图4. 邮件扫描流程



   \4. NavParser解析Excel电子表格

​       Python提供了多个能处理Excel电子表格的库，这里采用操作比较简单xlrd库对文件进行读取解析。操作步骤是：

​        1). 打开工作表文件

​        2). 获取表单（Sheet）列表名称或者索引

​        3). 根据名称或者索引取得当前表单

​        4). 逐行、逐列遍历Excel表格单元，查找所需要的内容

​       图5的代码片断就是解析文件的过程，扫描电子表格的每个单元，提取出单位净值、累计收益率等信息，然后返回一个字典结构，字典键就是产品名，数据就是提取到的净值、规模、收益率数据。提取数据的几个方法这里我都实现成私有方法（函数名前有2个下划线）。

​        

![img](http://img1.qdcypf.com/mmbiz_jpg/ANUAFkOeAz0IaeZHIbq1pqwWlRCDUVcBgtSCficfl2rQHszX5IY2csCCR62aQDPufeMWqzgAOyFDA3edVKWEiaog/640?wx_fmt=jpeg)

图5. 解析电子表格文件



   \5. ReportFiller模块

​      读完表后就开始填表。Python对Excel操作的库实现得比较拧巴，xlrd和xlwt两个模块一个只负责读，一个只负责写，并且，负责写的模块只支持新建写入文件，不能直接在原有文件上覆盖写入，以至于对报告进行每日增量更新比较麻烦。还好有人提供了一个xlutils库，实际上是在xlrd和xlwt上加了一层管道，把二者结合起来，但使用起来依然不是那么顺手。另外，这里面关于写入文件如何保持原有文件格式也需要一个技巧，在下一部分避坑指南中会详细介绍。最后，这些操作只能支持.xls格式的Excel文件，对于较新的.xlsx格式文件尚不支持。

​       ReportFiller使用xlrd和xlutils两个库结合实现原Excel表格文件的增量更新。如果要保留原来文件格式，需要在打开工作簿时候，设置参数formatting_info=True。xlutils会复制一个工作簿进行修改工作，修改完毕后将工作簿写回原来文件。图6是填写工作簿的代码片段。

​      从图中代码示例可见，使用xlrd.open_workbook()方法时候，参数formatting_info被设置为True。第54行是xlutils的拷贝操作，其内部本质还是调用了xlwt模块方法，因为xlwt模块不支持在原有Excel文件上直接修改，这个拷贝操作实际上是创建了一个新的工作簿。循环体中的操作是先从原工作簿中定位到需要填写的Excel单元行、列坐标，然后将要写入的内容填写到新拷贝的工作簿中。注意最后文件保存操作wrt_wb.save()是在with语句之外。这是因为with语句结束后才会释放资源（读写锁），关闭被打开的文件。这样保存操作才会成功将原报告文件覆盖，否则可能会引发无法写入已打开文件的错误。



![img](http://img1.qdcypf.com/mmbiz_jpg/ANUAFkOeAz0IaeZHIbq1pqwWlRCDUVcB3ibK1E3uKY3uJp9MvXV5vhorfJTq7Px0t7CsGd6PNuPN6yXaOD6I7ow/640?wx_fmt=jpeg)

图6. 填写Excel电子表格代码片段 



  \6. ReportSender模块

​      相比前面扫描邮件、读写电子表格，ReportSender模块功能就简单多了——创建一个电子邮件，把填写好的报告作为附件发送给相关人员。 创建邮件体使用email.mime类中提供的各种方法，然后使用smtplib模块，登录smtp服务器，将邮件发送出去即可。图7是整个邮件发送的核心代码。   



![img](http://img1.qdcypf.com/mmbiz_jpg/ANUAFkOeAz0IaeZHIbq1pqwWlRCDUVcBelS7NYWk6JVjjtic0dzdGuicOmEDCbIkursAn5uG4xOc3zuDicttXRrLQ/640?wx_fmt=jpeg)

图7. 使用Python的smtplib发送电子邮件



**四、技巧和避坑指南**

​      最开始编写这个程序，我的初衷就是实现一个提高工作效率的小工具，预计差不多半天到一天就可以完成，结果没想到中间还踩了几个坑，花了大概三天左右时间才初步完成，中间也学习到一些技巧，有了一些收获，这里把有意思的干货总结一下。

   \1. 如何优雅的跳出两层循环

​      在遍历扫描Excel表格时候，通常采用先按行再按列的双层循环结构，找到目标单元后，需要结束循环，进入下一步操作。如果在最内层循环调用break，只能中断最内层循环，外面一层还会继续。如果设置一个标识变量，比如loop_flag，中断内层循环后设置标识，外层再判断这个标识决定是否中断。感觉这样实现起来实在不够“优雅”或者“Pythonic”，有什么办法能够看起来更“舒适”一些呢？

​      我发现使用Python语言的一个特性可以实现，就是循环结束后的“else”语句。具体实现如图8代码所示。



![img](http://img1.qdcypf.com/mmbiz_jpg/ANUAFkOeAz0IaeZHIbq1pqwWlRCDUVcBMRjChIHibntS0sHrow98Cuhdzicaqz65ziccvf3ibicHDu1Io0SwkshvCmw/640?wx_fmt=jpeg)

图8. Python终结两层循环



​       第73行的“else”语句是Python特有的语法，这个else语句是距离它最近的一个循环如果正常结束就会被执行，如果被break中断掉，就不会执行。这样逻辑就很清晰了，如果70~72行的循环体一直循环到结束没有被中断，则73~74行语句会被执行，continue语句会跳过76行的break；如果70~72行的循环被72的break语句中断，则73~75行的语句不会被执行，76行的break会被执行，这样结果就是内层循环被中断，外层的同时也会被中断。

​      我发现Python这个循环体自带的“else”语句很有意思，第一次接触时候不太习惯甚至有点排斥，后面遇到一些应用场景后，感觉挺好用的，当然也有人批评这个特性让代码可读性变差。

  \2. 邮件编码解码

​      我们知道Email内容在传输时是经过编码的，最知名的就是base64编码，但其实MIME还定义一种QuotePrintable（简称QP）编码。由于只知道base64编码，以至于我遇到QP编码的邮件解码总是失败，开始还以为程序问题，后来找了一些经验资料才知道是解码方式不对。为此，邮件扫描程序实现了一个邮件解码方法，根据邮件提供的编码信息判断是base64还是QP编码，然后调用相应的解码器解出邮件原文。邮件解码方法如图9所示。

​    

![img](http://img1.qdcypf.com/mmbiz_png/ANUAFkOeAz0IaeZHIbq1pqwWlRCDUVcBc6RTCr3c4vIADMpDuUhrAsiaNWEANPHqgI4DdPCeNs8JyfwS3sbmMeA/640?wx_fmt=png)

图9. 处理邮件编码解码问题



  \3. Excel保留格式问题

​      使用xlwt或者xlutils填写Excel表格，会有一个恼人的问题：如果你的当前工作表拷贝自一个原工作表，在写入操作后，被写入单元的格式信息会丢失。解决办法是：在写入前先保存格式信息，然后进行写入单元格操作，再把保留的格式信息赋值给被写入的单元格。具体操作见图10。

![img](http://img1.qdcypf.com/mmbiz_png/ANUAFkOeAz0IaeZHIbq1pqwWlRCDUVcBzeXQ9EicSxQg14fmIbNHGiauo2lL9INqZYEZIianB7ibza6PIgDcyCia3nQ/640?wx_fmt=png)

图10. 如何对Excel单元进行写操作并保留原来格式



  \4. Excel文件保存问题

​      使用xlutils时候还遇到一个莫名其妙问题是，如果对Excel文件做了些许改动，比如打开文件手工修改一下格式等，修改完毕关闭文件，再用xlutils操作Excel写入后，保存文件会报错。报错原因也很无厘头，提示内容往往是：TypeError: save descriptor 'decode' requires a 'bytes' object but received a 'NoneType'。经过一番查询，貌似这是一个库文件的bug，应该是Unicode编码处理问题，需要Hack一下库文件，修改后即正常。可以参考网上内容https://github.com/python-excel/xlutils/issues/11   

   具体修正方式是:进入到Python安装目录下，找到Lib/site-packages/xlwt/UnicodeUtils.py文件，把里面的upack2()方法修改成图11的形式。



![img](http://img1.qdcypf.com/mmbiz_png/ANUAFkOeAz0IaeZHIbq1pqwWlRCDUVcBwd39KGrTAiaH1icewyBRd14ybnHXpBxZicYYtByWao66uIEf7k2XicY3ZQ/640?wx_fmt=png)

图11. 修正xlwt/UnicodeUtils.py中的Bug



**五、如何获得代码**

​      前文洋洋洒洒几千字介绍了很多内容，码农届有句话叫：“Talk is cheap, show me the code”。光有文字没有代码怎么能说明问题，这里告诉你如何获得参考代码实现——在github网站下载，地址：https://github.com/swankong/NavReporter 。代码中已经去掉了我工作相关的敏感信息，包括产品名，公司名，邮箱地址等。

​      由于这个程序较多依赖于我的工作环境，我又急于趁热打铁写成这篇文章，导致修改过后的代码大概不能直接成功运行，后面我会抽空花些时间维护一下，力争能做成一个例子，可以让github上下载的程序直接成功运行，这样就更有参考意义。还要强调一下，目前版本中的一些功能方法稍加修改就可以拿出来多带带使用，有类似需求的读者可自由参考借鉴。



**六、总结与展望**

​     本文介绍了如何使用Python实现自动邮件扫描和Excel读写，以达到办公自动化提高工作效率的目的**。**目前整个系统还只是完成了初步的功能实现，尚且存在一些缺陷。下一步的工作包括但不限于：

​        1). 完善程序中英文注释信息，提高可读性。

​        2). 完成一个去掉敏感信息的应用实例，实现github上传的程序可以直接运行。

​        3). 考虑修改程序功能，增加把数据写入关系数据库的功能，从而实现估值表扫描下载、读取解析、数据入库一整个流程自动化。

​       最后感谢互联网上众多技术爱好者发布的经验分享文章，让我在程序踩坑时候及时找到了解决方案。


  
