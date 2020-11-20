# rang-poi

**养成习惯,先赞后看!!!**
@[TOC](POI+EasyExcel学习笔记)

本片文章的项目GitHub地址:[https://github.com/haha143/rang-poi](https://github.com/haha143/rang-poi)

如果可以的话,欢迎大家star!!!!!

# 1.前言

相信大家都应该看到过下面的功能:

![在这里插入图片描述](https://img-blog.csdnimg.cn/20201120151018731.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

**`文件的导入导出`**:

这个功能主要就是帮助我们的用户能够快速的将数据导入到数据库中,不用在自己手动的一条一条的将数据新增到我们的数据库中.同时又能够方便我们能够将数据导出之后打印出来给领导们查看.不用非得带着电脑这里那里的跑.非常实用的功能.

功能好是好,但是这样的功能我们又应该怎么来开发呢,主要用到的技术又有那些呢?知道该用那些技术之后,我们又应该怎么来使用呢?相信大家肯定都有这样那样的困扰,正是因为大家有这样那样的困扰,所以就更加需要看看这篇文章了.这里我会用**案例+代码+源码**的方式带大家更好的学习这方面的知识.

文件的导入导出功能目前主要是两家独大,一个就是`Apache的POI`,另一家就是`阿里的EasyExcel`.这里两种技术我都会在下面的文章里面详细讲解.

# 2.POI:

## 2.1-POI介绍:

POI的全称是: **Poor Obfuscation Implementation** ,意思是可怜的模糊实现.说是可怜但是一点都不可怜.是由Apache公司用Java开发并且免费开源的一套Java Api.

它能够帮助我们简单快速的对Excel的数据进行读写的操作.他不仅支持Excel,同时也支持PowerPoint,Word等等,但是这两者我们就暂时不讲,我们主要需要了解的就是关于Excel的操作.

POI所需的依赖:

```java
<!--        xls03版本-->
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>3.9</version>
        </dependency>
<!--        xlsx07版本-->
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>3.9</version>
        </dependency>
<!--        日期格式化工具-->
        <dependency>
            <groupId>joda-time</groupId>
            <artifactId>joda-time</artifactId>
            <version>2.10.1</version>
        </dependency>
<!--        test-->
        <dependency>
            <groupId>junit</groupId>
            <artifactId>junit</artifactId>
            <version>4.12</version>
        </dependency>
```

## 2.2-03版Excel与07版Excel区别

在使用POI之前,我们需要先了解一下Excel的版本更替,这样能够方便我们更好的了解POI的使用.

这里面Excel主要就是有两类,分别是**Excel03版本**和**Excel07版本**

这两个版本之间主要有以下的差别:

- 两者数据量都是有限制的

  03版本行数最多只能到65536,列数最多只能到256

  03版本行数最多只能到1048576 ,列数最多只能到16384 

  03版本:

  行数限制:


![在这里插入图片描述](https://img-blog.csdnimg.cn/2020111711302048.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

  列数限制:

  9*29+22=256(简单的四则运算)

![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117113052892.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

  07版本:

  行数限制:


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117113110973.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

  列数限制:


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117113130903.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

- 两者的文件名后缀也不一样,03版本的后缀是xls,07版本的后缀是xlsx,既然两者的后缀不一样就说明操作两者的工具类肯定也就是不一样的,这一点我们会在下面的代码中着重体现,其次就是 **.xlsx文件比.xls的压缩率高，也就是相同数据量下，.xlsx的文件会小很多。** 

## 2.3-数据写入操作

知道上述两者的差异之后,才能更好的方便我们下面处理我们编写过程中可能遇到的bug.

其次在java中有一个非常重要的理念就是"**`万物皆对象`**",所以我们想要操作Excel表格的话,就要知道表格具体是由那些对象构成的.


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117113149319.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

具体分下来主要就是图中标注出来的几种对象:**工作簿,工作表,行,单元格**

了解完有上述对象之后,我们就通过一个简单的案例来帮助大家更好的 了解这个概念.

具体代码实现:

- 03版本-HSSFWorkbook:

```java
 @Test
    public void  testExcel03() throws Exception{
        //创建一个工作簿
        Workbook workbook=new HSSFWorkbook();
        //创建一张工作表
        Sheet sheet=workbook.createSheet("我是一个新表格");
        //创建一行即(1,1)的单元格
        Row row1=sheet.createRow(0);
        Cell cell11=row1.createCell(0);
        //往该单元格中填充数据
        cell11.setCellValue("姓名");
        //创建(1,2)单元格
        Cell cell12=row1.createCell(1);
        cell12.setCellValue("印某人");


        Row row2=sheet.createRow(1);
        Cell cell21=row2.createCell(0);
        cell21.setCellValue("注册日期");
        Cell cell22=row2.createCell(1);
        String time=new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);
        //创建文件流
        FileOutputStream fileOutputStream=new FileOutputStream(PATH+"登记表03.xls");
        //把文件流写入到工作簿中
        workbook.write(fileOutputStream);
        //关闭文件流
        fileOutputStream.close();
        System.out.println("文件生成成功");
    }
```

运行完代码之后我们就可以看到我们的文件夹下面就生成了**登记表03.xls**这样一个文件

![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117113209111.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

打开文件之后我们也能发现,数据的确已经写进来了.

![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117113228581.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

- 07版本-XSSFWorkbook:

```java
@Test
    public void testExcel07()throws Exception{
        //注意只有这里创建的对象是不一样的!!!!!
        Workbook workbook=new XSSFWorkbook();
        Sheet sheet=workbook.createSheet("我是一个新表格");
        Row row1=sheet.createRow(0);
        Cell cell11=row1.createCell(0);
        cell11.setCellValue("姓名");
        Cell cell12=row1.createCell(1);
        cell12.setCellValue("印某人");
        
        
        Row row2=sheet.createRow(1);
        Cell cell21=row2.createCell(0);
        cell21.setCellValue("注册日期");
        Cell cell22=row2.createCell(1);
        String time=new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        FileOutputStream fileOutputStream=new FileOutputStream(PATH+"登记表07.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("文件生成成功");
    }
```

可以看到在上面的代码中我们除了修改了创建的对象之外,其他的代码,我们都是没有修改的.本质操作都是一致的.

## 2.4-HSSFWorkbook,XSSFWorkbook,SXSSFWorkbook大数据量下写入速度对比

我们了解了基本的写入数据的流程之后,接下来我们测试一下,在大数据量的情况下,他们生成相应的文件需要多长的时间,看看他们两者的性能又是如何的.顺便我们也了解一下他们写入数据的整个流程.

- 03版本-HSSFWorkbook:

```java
@Test
    public void test03BigData()throws Exception{
        Long begin=System.currentTimeMillis();
        Workbook workbook=new HSSFWorkbook();
        Sheet sheet=workbook.createSheet();
        for(int rownum=0;rownum<65536;rownum++){
            Row row=sheet.createRow(rownum);
            for(int cellnum=0;cellnum<10;cellnum++){
                Cell cell=row.createCell(cellnum);
                cell.setCellValue(cellnum);
            }
        }
        FileOutputStream fileOutputStream=new FileOutputStream(PATH+"test03BigData.xls");
        workbook.write(fileOutputStream);
        System.out.println("文件生成完毕");
        Long end=System.currentTimeMillis();
        System.out.println("共用时:"+(double)(end-begin)/1000+"秒");

    }
```

这里我们运行完成之后可以看到一共运行了1.811秒,还是很快的

![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117113247161.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

但是就上我们上面所说的一样,03版本的只支持最多65536条数据,如果超过这个数据量的话,是会报这个错的: **Invalid row number (65536) outside allowable range (0..65535)** ,这里我们运行测试一下看一下报错:

![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117113303941.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

看完他的运行速度之后我们来看看,为什么HSSFWorkbook能够这么快就能将数据写入到文件中呢.

因为HSSFworkbook是直接将整个文件写入到内存中的,文件直接就能从内存中读到,所以使得整个写入的过程十分的快速.既然选择写入内存里面,那么就会出现一个问题那就是内存不够,直接就爆了,严重影响性能,所以可能是出于这个问题的考虑,03版本才会限制数据的条数(**后面部分是我自己猜的,嘤嘤嘤**)
![在这里插入图片描述](https://img-blog.csdnimg.cn/2020112018545472.png#pic_center)

- 07版本-XSSFWorkbook:

```java
 @Test
    public void test07BigData()throws Exception{
        Long begin=System.currentTimeMillis();
        Workbook workbook=new XSSFWorkbook();
        Sheet sheet=workbook.createSheet();
        for(int rownum=0;rownum<65537;rownum++){
            Row row=sheet.createRow(rownum);
            for(int cellnum=0;cellnum<10;cellnum++){
                Cell cell=row.createCell(cellnum);
                cell.setCellValue(cellnum);
            }
        }
        FileOutputStream fileOutputStream=new FileOutputStream(PATH+"test07BigData.xlsx");
        workbook.write(fileOutputStream);
        System.out.println("文件生成完毕");
        Long end=System.currentTimeMillis();
        System.out.println("共用时:"+(double)(end-begin)/1000+"秒");
    }
```

这里我们再来看一下XSSFWorkbook写入数据的速度:

同样的数据量,用时6.633秒

![1605581923483.png](https://img-blog.csdnimg.cn/20201117113326939.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)


我们再来看一下如果是10万条数据的话,看看时间会是多少:

用时10.013秒,时间还能接受,毕竟在10万条数据的情况下


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117113352623.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

既然这样我们也来分析一下XSSFWorkbook写入数据的流程,这里XSSFWorkbook和HSSFWorkbook一样,也是直接将数据写入内存中的,但是我们要知道因为XSSFWorkbook支持的数据量更多了,所以就必定会出现OOM即内存爆掉的情况,所以怎么办呢,这里我猜想的是,他是按照一定的量来将数据写入内存之中,就好比我是每10000条写入内存一次,那样的话,既能较快的写入数据,同时又能够支持比较大的数据量----这里也是我自己的猜想,感觉应该是这样.

大致可以通过下面的图来模拟:

![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117135132364.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

- 07版本进阶-SXSSFWorkbook

```java
 @Test
    public void test07BigDataS()throws Exception{
        Long begin=System.currentTimeMillis();
        Workbook workbook=new SXSSFWorkbook();
        Sheet sheet=workbook.createSheet();
        for(int rownum=0;rownum<65536;rownum++){
            Row row=sheet.createRow(rownum);
            for(int cellnum=0;cellnum<10;cellnum++){
                Cell cell=row.createCell(cellnum);
                cell.setCellValue(cellnum);
            }
        }
        FileOutputStream fileOutputStream=new FileOutputStream(PATH+"test07BigDataS.xlsx");
        workbook.write(fileOutputStream);
        System.out.println("文件生成完毕");
        Long end=System.currentTimeMillis();
        System.out.println("共用时:"+(double)(end-begin)/1000+"秒");
    }
```

大家看名字就知道这个SXSSFWorkbook其实就是XSSFWorkbook的加强版(**Super XSSFWorkbook**),他的优点比较明显,既能够支持写入大量的数据,同时写入数据的速度也是非常的快.

这里我们上来就直接测试10万条数据玩玩:


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117113413986.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

这速度跟闹着玩一样,10万条数据只要1.813秒,属实是牛逼

既然这样我们就更加要深挖一下,这玩意儿为啥这么快呢?

按照网上的说法,其实XSSFWorkbook写入数据的思路和XSSFWorkbook写入数据的思路差不多的,上面我们说过了XSSFWorkbook写入数据是每隔一个数据量进行输入,在已经向内存写入10000条数据后,程序就在进行等待,

等待着10000条数据写入文件之后,他才继续向内存里面写入数据.

SXSSFWorkbook的思路是这样,他一开始也是向`内存`里面写入数据,但是他有一个临界值默认是100.超过这个数据量之后的数据,他会自动在`磁盘`上创建一个临时文件,将数据写入该`文件`中,之后当内存中的数据写完之后就直接从临时文件中将数据拷贝过来,这样就大大的节省了时间,可以看到程序执行过程是没有断开的,是一直在执行的,意味着最耗时的部分一直在工作.所以才会使得SXSSFWorkbook既能写入大量的数据,同时又能够在非常快的时间内完成.

大致可以通过下面的图来模拟:

 ![img](https://img-blog.csdnimg.cn/20201117135116215.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)
其次就是SXSSFWorkbook与XSSFWorkbook有本质上的区别,这个我们可以通过他们引入的包名看出来:

 ![img](https://img-blog.csdnimg.cn/20201117150408488.png#pic_center) 

可以看到SXSSFWorkbook本质上是通过**流**来实现的,XSSFWorkbook则还是通过usermodel来实现的.显然流肯定是更快一点的.

并且这个临时文件并不是直接显示在项目路径下的一般都是存储在与该路径类似的路径下:**C:\Users\瓤瓤\AppData\Local\Temp**

这是我写入数据时生成的临时文件:

 文件名一般都是以POI开头![img](https://img-blog.csdnimg.cn/20201117150150484.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

 ## 2.5-POI数据读取操作

03版本-HSSFWorkbook:

```java
@Test
public void test03Read()throws Exception{
    //引入输入文件流
    FileInputStream fileInputStream=new FileInputStream(PATH+"test03BigData.xls");
    //创建工作簿
    Workbook workbook=new HSSFWorkbook(fileInputStream);
    //通过索引创建工作表
    Sheet sheet=workbook.getSheetAt(0);
    //通过索引获取行
    Row row=sheet.getRow(0);
    //通过索引获取单元格
    Cell cell=row.getCell(0);
    //打印单元格内容
    System.out.println(cell.getNumericCellValue());
}
```

这是最简单的读写操作流程.并且其中的工作表,行,单元格都是通过索引来获取,除了索引,POI还为我们提供了其他的获取方法,下面我们来详细说明一下.

获取工作表:


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117201853405.png#pic_center)

第一种就是直接通过工作表的表名来进行获取,第二种就是直接通过工作簿内工作表的索引来进行获取.

获取行就是只能通过索引来获取

剩下的就是获取单元格了:


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117201908175.png#pic_center)

第一种也是直接通过索引来进行获取,第二种只不是多加了一层的判断语句,这个我们可以点进源码里面看一下:


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201117201908302.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

主要有这三个

- RETURN_NULL_AND_BLANK

  英文解释： **Missing cells are returned as null, Blank cells are returned as normal** 

  缺失的单元格会返回为空，空的单元格就正常返回即可。

- RETURN_BLANK_AS_NULL

   英文解释： **Missing cells are returned as null, as are blank cells** 

​       缺失的单元格返回为空，空的单元格也是如此。

- RETURN_BLANK_AS_NULL

  英文解释： **A new, blank cell is created for missing cells. Blank cells are returned as normal** 

  缺失的单元格不仅返回为空，同时还将为这个缺失的单元格创建一个新的单元格。空的单元格就正常返回即可。

其实这三种概念的理念差不多，基本上主要都是用来处理如果出现缺失的单元格情况时，可能会影响后续数据的读写操作。

07版本-XSSFWorkbook:

```java
@Test
    public void test07Read()throws Exception{
        FileInputStream fileInputStream=new FileInputStream(PATH+"test07BigData.xlsx");
        Workbook workbook=new XSSFWorkbook(fileInputStream);
        Sheet sheet=workbook.getSheetAt(0);
        Row row=sheet.getRow(0);
        Cell cell[]=new Cell[4];
        cell[0]=row.getCell(0);
        cell[1]=row.getCell(1,Row.RETURN_NULL_AND_BLANK);
        cell[2]=row.getCell(2);
        cell[3]=row.getCell(3);
        for(int i=0;i<4;i++){
            System.out.println(cell[i]);
        }
    }
```

可以看到我们只是简单的修改了一下对象,其他的操作我们都是没有改的,所以我们在编写的时候,只需要注意我们版本对应的对象就行了.

## 2.6-POI读取不同数据类型的数据

表格数据:


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201120151113598.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)


这里我们已经将我们平常能够遇到的数据类型全部都包含到了.

接下来我们通过这段代码进行测试:

```java
@Test
    public void testMultipleTypeRead()throws Exception{
        FileInputStream fileInputStream=new FileInputStream(PATH+"test.xls");
        Workbook workbook=new HSSFWorkbook(fileInputStream);
        //获取表格的列名
        Sheet sheet=workbook.getSheetAt(0);
        Row rowTitle=sheet.getRow(0);
        int cellNum=rowTitle.getLastCellNum();
        if(rowTitle!=null){
            for(int i=0;i<cellNum;i++){
                Cell cell=rowTitle.getCell(i);
                int cellType=cell.getCellType();
                if(cell!=null){
                    System.out.print(cell+"-"+cellType+" | ");
                }
            }
        }
        System.out.println();
        //获取表格的数据部分
        int RowNum=sheet.getLastRowNum();
        for(int i=1;i<=RowNum;i++){
            Row rowData=sheet.getRow(i);
            if(rowData!=null){
                int cellnum=rowData.getLastCellNum();
                for(int j=0;j<cellnum;j++){
                    Cell cell=rowData.getCell(j);
                    int cellType=cell.getCellType();
                    if(cell!=null){
                        //根据单元格数据类型进行相应的数据输出
                        switch (cellType){
                            //数字类型数据
                            case HSSFCell.CELL_TYPE_NUMERIC:
                                System.out.print(cell.getNumericCellValue()+"-"+cellType+" | ");
                                continue;
                            //字符串类型数据
                            case HSSFCell.CELL_TYPE_STRING:
                                System.out.print(cell.getStringCellValue()+"-"+cellType+" | ");
                                continue;
                            //公式类型
                            case HSSFCell.CELL_TYPE_FORMULA:
                                System.out.print("null"+"-"+cellType+" | ");
                                continue;
                            //空单元格
                            case HSSFCell.CELL_TYPE_BLANK:
                                System.out.print(cell.getStringCellValue()+"-"+cellType+" | ");
                                continue;
                            //布尔值类型
                            case HSSFCell.CELL_TYPE_BOOLEAN:
                                System.out.print(cell.getBooleanCellValue()+"-"+cellType+" | ");
                                continue;
                            //错误单元格
                            case HSSFCell.CELL_TYPE_ERROR:
                                System.out.print(cell.getErrorCellValue()+"-"+cellType+" | ");
                                continue;
                        }
                    }
                }
                System.out.println();
            }
        }
    }
```

这里我们可以看到能够输出下面的结果:


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201120151204498.png#pic_center)

其中上面的单元格类型变量,我们既可以通过直接的0,1,2....来定义,同时也能够直接通过HSSFCell的变量值来直接定义.

这里为了方便大家更好的理解,我们点进源码查看一下:

我们进入HSSFCell之后并没有看到我们想要的变量名:


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201120151233611.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

但是我们看到HSSFCell他是实现了Cell这个接口的,所以不出意外,这些变量应该就是在Cell里面定义,所以我们再点进Cell里面看.发现的确就是如我们想的一样:


![在这里插入图片描述](https://img-blog.csdnimg.cn/202011201513037.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

并且他们的返回值都是int类型的,所以这就行号解释了为什么能够直接调用这些变量了.

## 2.7-POI计算公式

这里我们在之前的test.xls文件里面为一个单元格增加了一个公式:


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201120151332453.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

接下来我们通过下面的代码将公式以及公式计算的结果读取出来:

```java
@Test
    public void testFORMULA()throws Exception{
        FileInputStream fileInputStream=new FileInputStream(PATH+"test.xls");
        Workbook workbook=new HSSFWorkbook(fileInputStream);
        //获取到包含公式的单元格
        Sheet sheet=workbook.getSheetAt(0);
        Row row=sheet.getRow(3);
        Cell cell=row.getCell(7);
        //读取计算的公式
        FormulaEvaluator formulaEvaluator=new HSSFFormulaEvaluator((HSSFWorkbook) workbook);
        int cellType=cell.getCellType();
        switch (cellType){
            //单元格的类型是公式类型
            case HSSFCell.CELL_TYPE_FORMULA:
                //公式内容
                String formula=cell.getCellFormula();
                System.out.println(formula);
                //执行公式之后,单元格内的值
                CellValue evaluate=formulaEvaluator.evaluate(cell);
                String cellValue=evaluate.formatAsString();
                System.out.println(cellValue);
                break;
        }
```

接着我们来运行一下,看看结果吧:


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201120151402362.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

可以看到输出的结果和我们在Excel里面看到的结果是一样的.

到这里我们关于POI的操作基本就已经结束了,接下来我们就主要了解一下EsayExcel.

# 3.EsayExcel:

真的是没有对比就没有伤害，在使用POI的过程中，感觉整个的流程还是比较简单的,毕竟就和我们平常写Excel表格的步骤是一样的,但是在真正使用了EasyExcel之后才发现,POI真的是弱爆了,并且在POI中我们需要使用到大量的for循环,这样会严重影响我们程序的性能,但是EasyExcel就已经帮我们优化好了,使得整个程序的性能一直处于十分强悍的状态.

并且就如同我们上面分析过的一样,POI本质上主要是在内存中进行数据的读写,但是在EasyExcel中就不一样了,他是直接将大部分的工作直接转移到了硬盘上这样就能大大减少我们内存的使用,性能能够得到大幅度的提升.对比如下图所示:

 ![POI与easyExcel对比图](https://img-blog.csdnimg.cn/20190606114105648.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQwNTIxNjU2,size_16,color_FFFFFF,t_70) 

## 3.1-EasyExcel介绍:

 EasyExcel是一个基于Java的简单、省内存的读写Excel的开源项目。在尽可能节约内存的情况下支持读写百M的Excel。 

EasyExcel的GitHub地址:[https://github.com/alibaba/easyexcel](https://github.com/alibaba/easyexcel)

EasyExcel的官方文档:[https://www.yuque.com/easyexcel/doc/easyexcel](https://www.yuque.com/easyexcel/doc/easyexcel)

其实简单总结一下EasyExcel的特点就是一个字:快.

EasyExcel所需的依赖:

```java
<dependency>
    <groupId>com.alibaba</groupId>
    <artifactId>easyexcel</artifactId>
    <version>2.2.6</version>
</dependency>
```

我们在引入EasyExcel依赖之后,我们需要注意下面的问题,因为EasyExcel里面已经集成了很多的依赖,并且里面就包含了POI的依赖:



![在这里插入图片描述](https://img-blog.csdnimg.cn/20201120151502931.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)


所以我们需要将我们之前引入的POI的依赖注释掉,否则会出现依赖的重复.

## 3.2-EasyExcel数据写入操作

  首先我们需要创建一个实体类.用来映射到我们在Excel中将要填充的对象

```java
import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;
import java.util.Date;
@Data
public class DemoData {
    @ExcelProperty("字符串标题")
    private String string;
    @ExcelProperty("日期标题")
    private Date date;
    @ExcelProperty("数字标题")
    private Double doubleData;
    /**
     * 忽略这个字段
     */
    @ExcelIgnore
    private String ignore;
}
```

并且EasyExcel还为我们提供了一些注解,方便我们的工作

@ExcelProperty(""):用来标注Excel中字段的标题

@ExcelIgnore:用来表示该字段忽略,不用添加到Excel中

```java
public class TestEasyExcel {
    String PATH="D:/software/IDEA/projects/rang-poi/";
    //填充我们即将写入Excel中的数据
    private List<DemoData> data() {
        List<DemoData> list = new ArrayList<DemoData>();
        for (int i = 0; i < 10; i++) {
            DemoData data = new DemoData();
            data.setString("字符串" + i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }

    /**
     * 最简单的写
     * <p>1. 创建excel对应的实体对象 参照{@link DemoData}
     * <p>2. 直接写即可
     */
    @Test
    public void simpleWrite() {
        // 写法1
        //创建文件名
        String fileName = PATH+ "easyexcel.xlsx";
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        // 如果这里想使用03 则 传入excelType参数即可
        EasyExcel.write(fileName, DemoData.class).sheet("模板").doWrite(data());
        
        
//        // 写法2
//        fileName = PATH+ "easyexcel.xlsx";
//        // 这里 需要指定写用哪个class去写
//        ExcelWriter excelWriter = null;
//        try {
//            excelWriter = EasyExcel.write(fileName, DemoData.class).build();
//            WriteSheet writeSheet = EasyExcel.writerSheet("模板").build();
//            excelWriter.write(data(), writeSheet);
//        } finally {
//            // 千万别忘记finish 会帮忙关闭流
//            if (excelWriter != null) {
//                excelWriter.finish();
//            }
//        }
    }
}
```

这样我们的数据写入就完成了,运行代码之后我们就可以看到已经在我们的项目路径下生成了easyexcel文件了

打开之后


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201120184930969.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

数据也的确已经插入进来了

上面的代码中有两段执行数据写入的方法,第一段代码就是直接将数据写入到文件中,第二段代码就类似于POI中的通过for循环将数据一条一条的写入进去,显然第二种方法效率较低,推荐使用第一种.这里对比POI之后,我们可以发现EasyExcel极大的降低了代码量.

## 3.3-EasyExcel数据读取操作

首先我们需要创建一个监听器:

```java
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

// 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
public class DemoDataListener extends AnalysisEventListener<DemoData> {
    private static final Logger LOGGER = LoggerFactory.getLogger(DemoDataListener.class);
    /**
     * 每隔5条存储数据库，实际使用中可以3000条，然后清理list ，方便内存回收
     */
    private static final int BATCH_COUNT = 5;
    List<DemoData> list = new ArrayList<DemoData>();
    /**
     * 假设这个是一个DAO，当然有业务逻辑这个也可以是一个service。当然如果不用存储这个对象没用。
     */
    private DemoDAO demoDAO;
    public DemoDataListener() {
        // 这里是demo，所以随便new一个。实际使用如果到了spring,请使用下面的有参构造函数
        demoDAO = new DemoDAO();
    }
    /**
     * 如果使用了spring,请使用这个构造方法。每次创建Listener的时候需要把spring管理的类传进来
     *
     * @param demoDAO
     */
    public DemoDataListener(DemoDAO demoDAO) {
        this.demoDAO = demoDAO;
    }
    /**
     * 这个每一条数据解析都会来调用
     *
     * @param data
     *            one row value. Is is same as {@link AnalysisContext#readRowHolder()}
     * @param context
     */
    @Override
    public void invoke(DemoData data, AnalysisContext context) {
        System.out.println(JSON.toJSONString(data));
        LOGGER.info("解析到一条数据:{}", JSON.toJSONString(data));
        list.add(data);
        // 达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
        if (list.size() >= BATCH_COUNT) {
            saveData();
            // 存储完成清理 list
            list.clear();
        }
    }
    /**
     * 所有数据解析完成了 都会来调用
     *
     * @param context
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        // 这里也要保存数据，确保最后遗留的数据也存储到数据库
        saveData();
        LOGGER.info("所有数据解析完成！");
    }
    /**
     * 加上存储数据库
     */
    private void saveData() {
        LOGGER.info("{}条数据，开始存储数据库！", list.size());
        demoDAO.save(list);
        LOGGER.info("存储数据库成功！");
    }
}

```

之后我们需要根据自己的需要创建一个DAO功能其实就类似于我们的service层,可以在这里面定义我们后来可能加入的与数据库的相关操作的方法

```java
/**
 * 假设这个是你的DAO存储。当然还要这个类让spring管理，当然你不用需要存储，也不需要这个类。
 **/
public class DemoDAO {
    public void save(List<DemoData> list) {
        // 如果是mybatis,尽量别直接调用多次insert,自己写一个mapper里面新增一个方法batchInsert,所有数据一次性插入
    }
}
```

创建完成之后我们的功能基本就可以了,之后就可以进行测试了:

```java
/**
     * 最简单的读
     * <p>1. 创建excel对应的实体对象 参照{@link DemoData}
     * <p>2. 由于默认一行行的读取excel，所以需要创建excel一行一行的回调监听器，参照{@link DemoDataListener}
     * <p>3. 直接读即可
     */
    @Test
    public void simpleRead() {
        // 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
        // 写法1：
        String fileName = PATH+ "easyexcel.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();

//        // 写法2：
//        String fileName = PATH+ "easyexcel.xlsx";
//        ExcelReader excelReader = null;
//        try {
//            excelReader = EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).build();
//            ReadSheet readSheet = EasyExcel.readSheet(0).build();
//            excelReader.read(readSheet);
//        } finally {
//            if (excelReader != null) {
//                // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
//                excelReader.finish();
//            }
//        }
    }
```

运行完成之后我们看到这样的结果:


![在这里插入图片描述](https://img-blog.csdnimg.cn/20201120185002569.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)

就说明的确是已经将数据读出来了.

这里其实和上面数据写入是一样的,同样也有两个方法.同样的第二个也是类似于for循环,循环遍历数据,所以 效率比较慢,还是建议第一种方法.

-----------------------------------------------------------------------------------------------------------------------------------------------------------

到这里POI和EasyExcel就已经介绍完成了,码字不易,如果觉得对你有帮助的话,可以关注我的公众号,新人up需要你的支持!!!

![在这里插入图片描述](https://img-blog.csdnimg.cn/20201120185026416.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L2xvdmVseV9fUlI=,size_16,color_FFFFFF,t_70#pic_center)


**不点在看,你也好看,**

**点了在看,你更好看!!**

