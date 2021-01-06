## QLExcelDemo

> OC中使用 [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter) 生成和导出excel文件的简单Demo
> 
> 相关文档地址：[https://libxlsxwriter.github.io/](https://libxlsxwriter.github.io/)，
> 所有函数、属性的使用及参数的含义里面都有详细介绍。这里只大致介绍一下常用的属性设置。

### 一、添加 `libxlsxwriter` 库
 - 方式1：使用cocoapods导入libxlsxwriter。
 - 方式2：可以直接拖入相关库（建议直接拖入demo里面的库，直接从github上下载会有些问题，此库拉取时间为2021年01月05日版本，也可以先用cocopods进行拉取后再找到源文件拖入正式项目）

### 二、使用方式

#### 1. 生成xlsx
 - 导入头文件

 ```
 #import "xlsxwriter.h"
 ```
 
 - 创建保存文件的路径：
 
 ```
 NSString *documentDirectory = [NSSearchPathForDirectoriesInDomains(NSDocumentDirectory, NSUserDomainMask, YES) objectAtIndex:0];
NSLog(@"documentDirectory === %@",documentDirectory);
 ```
 
 - 创建表格workbook，和工作表worksheet：

	```
	// 创建新xlsx文件，路径需要转成c字符串
    lxw_workbook  *workbook  = workbook_new([filename UTF8String]);
    // 创建sheet
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);
	```

 - 保存生成文件
 
	```
	//保存
    workbook_close(workbook);
	```

此时生成文件 name.xlsx ，但是没有写入内容，所以是个空的xlsx文件

#### 2. 设置文本格式

```
// 创建格式titleformat
lxw_format *titleformat = workbook_add_format(workbook);
// 字体加粗
format_set_bold(titleformat);
// 字体尺寸
format_set_font_size(titleformat, 20);
// 字体颜色
format_set_font_color(columnTitleformat,0x696969);
// 内容垂直居中
format_set_align(titleformat, LXW_ALIGN_VERTICAL_CENTER);
// 内容水平居中
format_set_align(titleformat, LXW_ALIGN_CENTER);
// 边框（四周）：中宽边框
format_set_border(columnDetailformat, LXW_BORDER_MEDIUM);  
// 右边框：双线边框
format_set_right(columnDetailformat, LXW_BORDER_DOUBLE);
// 左边框：双线边框
format_set_left(columnDetailformat, LXW_BORDER_DOUBLE);
// 下边框：双线边框
format_set_bottom(columnDetailformat, LXW_BORDER_DOUBLE);
    
// 合并单元格。0行0列 到 0行8列 合并为一行，并设定内容为"核销结算表"
worksheet_merge_range(worksheet, 0, 0,0, 8, [@"（文旅惠民消费券《核销结算表》" cStringUsingEncoding:NSUTF8StringEncoding], titleformat);
    
// 数字格式
format_set_num_format(formatNum, "#,##0.00");
```

设置行高、列宽

```
// 第1行的高度为30，并将格式titleformat应用到该行上。注意 行的高度和列表的宽度的值单位不一样，此30非彼30
worksheet_set_row(worksheet, 1, 30, titleformat);
 
 第0列到第8列的宽度为30。注意 此30非彼30
worksheet_set_column(worksheet, 0, 8, 30.0, NULL);
```

#### 3. 写入文本

```
// 将"填报日期：" 写入到1行7列，并按照columnTitleformat格式，可为NULL
worksheet_write_string(worksheet, 1, 7, [@"填报日期：" UTF8String], columnTitleformat);
// 写入数字格式文本
worksheet_write_number(worksheet, 3+i, 8, [model[@"payment"] doubleValue], formatNum);
```
#### 4. 使用数学公式计算

```
NSString *sumStr = [NSString stringWithFormat:@"=SUM(I4:I%lu)",3+self.accountDataArray.count];
// 采用数学公式计算
worksheet_write_formula(worksheet, rowA, 8, [sumStr cStringUsingEncoding:NSUTF8StringEncoding], formatNum);
```

### 三、部分参数解释

> [格式文档地址](http://libxlsxwriter.github.io/working_with_formats.html)

格式              | 说明             | 文档
:--------------  | :-------------: | -----------:
`format_set_align` | 设置对齐方式     | [文档](http://libxlsxwriter.github.io/format_8h.html#a189c83d1f21b01937f1f730720c33d13)
`format_set_border` | 设置边框        | [文档](http://libxlsxwriter.github.io/format_8h.html#a9cf7a28a6e8014cb98dff27415e2b1ca)
`format_set_num_format` | 设置数字格式  | [文档](http://libxlsxwriter.github.io/format_8h.html#af77bbd0003344cb16d455c7fb709e16c)
`worksheet_write_formula` | 数学公式   | [文档](http://libxlsxwriter.github.io/worksheet_8h.html#ae57117f04c82bef29805ec3eabc219bb)
`worksheet_freeze_panes` | 标题栏固定    | --
`worksheet_write_string` | 写入文本     | -- 
`worksheet_write_number` | 写入数字     | --
`worksheet_merge_range`  | 合并行、列   | --

### 四、备注

- 1. 在使用 `worksheet_write_formula` 时，首先要保证写入时使用 `worksheet_write_number` ，并且在预览时此值会为默认的 `0`，不过在实际打开此excel时会直接计算此公式得值，因为已经将公式在excel文件中写好了。
- 2. 在使用 `"=SUM(A1:A2)"`时，要将 `1，2`等行数 替换成 在excel打开时的行数或列数。

--

# BY QiuFairy
