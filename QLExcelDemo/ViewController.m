//
//  ViewController.m
//  QLExcelDemo
//
//  Created by qiu on 2021/1/5.
//  Copyright © 2021 qiu. All rights reserved.
//

#import "ViewController.h"
#import <WebKit/WebKit.h>
#import "xlsxwriter.h"

#define APP_SCREEN_HEIGHT [UIScreen mainScreen].bounds.size.height
#define NAV_HEIGHT 88.0f
#define BOTTOM_SAFE_AREA 39.0f
#define APP_SCREEN_WIDTH [UIScreen mainScreen].bounds.size.width

@interface ViewController () <WKNavigationDelegate, WKUIDelegate, UIDocumentInteractionControllerDelegate>

@property (nonatomic,strong) NSMutableArray *accountDataArray;
@property (nonatomic,strong) UIButton * exportAccountBtn;
@property (nonatomic,strong)UIDocumentInteractionController * document;
@property(nonatomic,assign) NSInteger startYear;
@property(nonatomic,assign) NSInteger startMonth;
@property(nonatomic,assign) NSInteger startDay;
@property(nonatomic,assign) NSInteger endYear;
@property(nonatomic,assign) NSInteger endMonth;
@property(nonatomic,assign) NSInteger endDay;

@property (nonatomic, strong) WKWebView *testWebView;

@end

@implementation ViewController
-(NSMutableArray *)accountDataArray{
    if(!_accountDataArray){
        _accountDataArray = [NSMutableArray array];
    }
    return _accountDataArray;
}

- (void)viewDidLoad {
    [super viewDidLoad];
    
    NSArray *arr = @[@{@"discountMoney":@"10",@"payment":@"11",@"companyName":@"12",@"accountNo":@"13",@"orderNo":@"14",@"price":@"15",@"allInPayOrderNo":@"16",@"discountMoney":@"17",@"payment":@"18"},
                    @{@"discountMoney":@"20",@"payment":@"21",@"companyName":@"22",@"accountNo":@"23",@"orderNo":@"24",@"price":@"25",@"allInPayOrderNo":@"26",@"discountMoney":@"27",@"payment":@"28"},
                    @{@"discountMoney":@"30",@"payment":@"31",@"companyName":@"32",@"accountNo":@"33",@"orderNo":@"34",@"price":@"35",@"allInPayOrderNo":@"36",@"discountMoney":@"37",@"payment":@"38"}];
    self.accountDataArray = [NSMutableArray arrayWithArray:arr];
    
    [self exportAccountUI];
    [self netRequest];
}
-(void)exportAccountUI{
    self.exportAccountBtn = [[UIButton alloc]initWithFrame:CGRectMake(0,APP_SCREEN_HEIGHT - BOTTOM_SAFE_AREA - 40 ,APP_SCREEN_WIDTH,40)];
    self.exportAccountBtn.backgroundColor = [UIColor lightGrayColor];
    [self.exportAccountBtn setTitle:@"导出Excel" forState:UIControlStateNormal];
    [self.exportAccountBtn setTitleColor:[UIColor redColor] forState:UIControlStateNormal];
    self.exportAccountBtn.titleLabel.font  = [UIFont systemFontOfSize:14];
    [self.exportAccountBtn addTarget:self action:@selector(exportAccountClick:) forControlEvents:UIControlEventTouchUpInside];
    [self.view addSubview:self.exportAccountBtn];
}
-(void)netRequest{
    
    // ....
    // 数据请求，拿到数据源
    
    // 指定路径下创建excel
    [self creationExcel];
}
#pragma mark - 创建excel
// 创建excel
-(void)creationExcel{
    // 获取路径
    NSString *documentDirectory = [NSSearchPathForDirectoriesInDomains(NSDocumentDirectory, NSUserDomainMask, YES) objectAtIndex:0];
    NSLog(@"documentDirectory === %@",documentDirectory);
    NSString *xlsxName = [NSString stringWithFormat:@"%ld年%ld月%ld日-%ld年%ld月%ld日对账单.xlsx",self.startYear,self.startMonth,self.startDay,self.endYear,self.endMonth,self.endDay];
    // 拼接 cache路径 + 名为xlsxName的文件   注意：该路径下的文件名一定要和后续导出的文件名保持一致，否则后续执行导出操作时，读取不到该文件
    NSString *filename = [documentDirectory stringByAppendingPathComponent:xlsxName];
    
    // 创建新xlsx文件，路径需要转成c字符串
    lxw_workbook  *workbook  = workbook_new([filename UTF8String]);
    // 创建sheet
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);
    
    // 格式1
    lxw_format *titleformat = workbook_add_format(workbook);
    // 字体加粗
    format_set_bold(titleformat);
    // 字体尺寸
    format_set_font_size(titleformat, 20);
    // 内容垂直居中
    format_set_align(titleformat, LXW_ALIGN_VERTICAL_CENTER);
    // 内容水平居中
    format_set_align(titleformat, LXW_ALIGN_CENTER);
    // 合并单元格。0行0列到0行8列合并为一行，并设定内容为"文旅惠民消费券《核销结算表》"
    worksheet_merge_range(worksheet, 0, 0,0, 8, [@"（文旅惠民消费券《核销结算表》" cStringUsingEncoding:NSUTF8StringEncoding], titleformat);
    
    // 格式2
    lxw_format *columnTitleformat = workbook_add_format(workbook);
    // 内容垂直居中
    format_set_align(columnTitleformat, LXW_ALIGN_VERTICAL_CENTER);
    // 内容水平居中
    format_set_align(columnTitleformat, LXW_ALIGN_CENTER);
    // 字体尺寸
    format_set_font_size(columnTitleformat, 17);
    // 字体颜色
    format_set_font_color(columnTitleformat,0x696969);
    // 字体加粗
    format_set_bold(columnTitleformat);
    
    // 将"统计周期：" 的总价写入到1行5列
    worksheet_write_string(worksheet, 1, 5, [@"统计周期：" UTF8String], columnTitleformat);
    // 将"填报日期：" 的总价写入到1行7列
    worksheet_write_string(worksheet, 1, 7, [@"填报日期：" UTF8String], columnTitleformat);
    // 第1行的高度为20，并将格式2应用到该行上。注意 行的高度和列表的宽度的值单位不一样，此50非彼50
    worksheet_set_row(worksheet, 1, 30, columnTitleformat);
    
    // 第0列到第8列的宽度为30。注意 此30非彼30
    worksheet_set_column(worksheet, 0, 8, 30.0, NULL);
    // 第2行的高度为30，并将格式2应用到该行上。注意 行的高度和列表的宽度的值单位不一样，此30非彼30
    worksheet_set_row(worksheet, 2, 30, columnTitleformat);
    
    // 将"序号"写入到2行0列
    worksheet_write_string(worksheet, 2, 0, [@"序号" UTF8String], columnTitleformat);
    // 将"企业名称"写入到2行1列
    worksheet_write_string(worksheet, 2, 1, [@"企业名称" UTF8String], columnTitleformat);
    // 将"商家账号"写入到2行2列
    worksheet_write_string(worksheet, 2, 2, [@"商家账号" UTF8String], columnTitleformat);
    // 将"消费日期"写入到2行3列
    worksheet_write_string(worksheet, 2, 3, [@"消费日期" UTF8String], columnTitleformat);
    // 将"销售订单号"写入到2行4列
    worksheet_write_string(worksheet, 2, 4, [@"销售订单号" UTF8String], columnTitleformat);
    // 将"销售金额"写入到2行5列
    worksheet_write_string(worksheet, 2, 5, [@"销售金额" UTF8String], columnTitleformat);
    // 将"通联订单号"写入到2行6列
    worksheet_write_string(worksheet, 2, 6, [@"通联订单号" UTF8String], columnTitleformat);
    // 将"优惠（券）"写入到2行7列
    worksheet_write_string(worksheet, 2, 7, [@"优惠（券）" UTF8String], columnTitleformat);
    // 将"实付金额"写入到2行8列
    worksheet_write_string(worksheet, 2, 8, [@"实付金额" UTF8String], columnTitleformat);
    
    int  rowA = 3 + (int)self.accountDataArray.count;
    // 合并单元格。rowA行0列 到 rowA行5列合并为一行，并设定内容为 "合计："
    worksheet_merge_range(worksheet, rowA,  0 ,rowA, 5, [@"合计：" cStringUsingEncoding:NSUTF8StringEncoding], columnTitleformat);
    // 第rowA行的高度为50。注意 行的高度和列表的宽度的值单位不一样，此50非彼50
    worksheet_set_row(worksheet, rowA, 50, columnTitleformat);

    
    float discountPrice = 0.00;
    float realPrice = 0.00;
    // 统计 优惠（券）所在列的总价 和 实付金额 所在列的总价
    for (int i = 0; i < self.accountDataArray.count; i++) {
        NSDictionary *model = self.accountDataArray[i];
        discountPrice = discountPrice + [model[@"discountMoney"] floatValue];
        realPrice = realPrice + [model[@"payment"] floatValue];
    }
    
    lxw_format *formatNum = workbook_add_format(workbook);
    // 数字格式
    format_set_num_format(formatNum, "#,##0.00");
    // 内容垂直居中
    format_set_align(formatNum, LXW_ALIGN_VERTICAL_CENTER);
    // 内容水平居中
    format_set_align(formatNum, LXW_ALIGN_CENTER);
    // 字体尺寸
    format_set_font_size(formatNum, 17);
    // 字体颜色
    format_set_font_color(formatNum,0x696969);
    // 字体加粗
    format_set_bold(formatNum);
    
    format_set_border(formatNum, LXW_BORDER_MEDIUM);// 边框（四周）：中宽边框
    format_set_right(formatNum, LXW_BORDER_DOUBLE);// 右边框：双线边框
    format_set_left(formatNum, LXW_BORDER_DOUBLE);// 左边框：双线边框
    format_set_bottom(formatNum, LXW_BORDER_DOUBLE);// 下边框：双线边框
    
    // 将优惠（券）所在列的总价写入到rowA行7列
    worksheet_write_string(worksheet, rowA, 7, [[NSString stringWithFormat:@"%.2lf",discountPrice] UTF8String], formatNum);
    // 这个计算将用数学公式替换
    // 将实付金额 所在列的总价写入到rowA行8列
//    worksheet_write_string(worksheet, rowA, 8, [[NSString stringWithFormat:@"%.2lf",realPrice] UTF8String], columnTitleformat);
    
    
    int  rowB = 4 + (int)self.accountDataArray.count;
    // 合并单元格。rowB行0列 到 rowB行1列合并为一行，并设定内容为 "负责人："
    worksheet_merge_range(worksheet, rowB,  0 ,rowB, 1, [@"负责人：" cStringUsingEncoding:NSUTF8StringEncoding], columnTitleformat);
    // 合并单元格。rowB行3列 到 rowB行4列合并为一行，并设定内容为 "填表人："
    worksheet_merge_range(worksheet, rowB,  3 ,rowB, 4, [@"填表人：" cStringUsingEncoding:NSUTF8StringEncoding], columnTitleformat);
    // 合并单元格。rowB行6列 到 rowB行7列合并为一行，并设定内容为 "单位：（盖章）"
    worksheet_merge_range(worksheet, rowB,  6 ,rowB, 7, [@"单位：（盖章）" cStringUsingEncoding:NSUTF8StringEncoding], columnTitleformat);
    // 第rowB行的高度为50，并将格式2应用到该行上。注意 行的高度和列表的宽度的值单位不一样，此50非彼50
    worksheet_set_row(worksheet, rowB, 50, columnTitleformat);
    
    // 格式3
    lxw_format *markformat = workbook_add_format(workbook);
    // 内容垂直居中
    format_set_align(markformat, LXW_ALIGN_VERTICAL_CENTER);
    // 内容水平居中
    format_set_align(markformat, LXW_ALIGN_RIGHT);
    // 字体尺寸
    format_set_font_size(markformat, 17);
    // 字体颜色
    format_set_font_color(markformat,0xFF0000);
    int  rowC = 5 + (int)self.accountDataArray.count;
    // 合并单元格。rowC行0列 到 rowB行8列合并为一行，并设定内容为 "备注1、使用文化和旅游惠民消费券的订单，以通联订单号为唯一标识逐笔填写。"
    worksheet_merge_range(worksheet, rowC,  0 ,rowC, 8, [@"" cStringUsingEncoding:NSUTF8StringEncoding], markformat);
    // 第rowC行的高度为50，并将格式3应用到该行上。注意 行的高度和列表的宽度的值单位不一样，此30非彼30
    worksheet_set_row(worksheet,  rowC, 30, markformat);
    
    // 合并单元格。rowC行0列 到 rowB行8列合并为一行，并设定内容为 "备注1、使用文化和旅游惠民消费券的订单，以通联订单号为唯一标识逐笔填写。"
    worksheet_merge_range(worksheet, rowC+1,  0 ,rowC+1, 8, [@"备注1、使用文化和旅游惠民消费券的订单，以通联订单号为唯一标识逐笔填写。" cStringUsingEncoding:NSUTF8StringEncoding], markformat);
    // 第rowC行的高度为50，并将格式3应用到该行上。注意 行的高度和列表的宽度的值单位不一样，此30非彼30
    worksheet_set_row(worksheet,  rowC+1, 30, markformat);
    
    // 格式4
    lxw_format *columnDetailformat = workbook_add_format(workbook);
    // 内容垂直居中
    format_set_align(columnDetailformat, LXW_ALIGN_VERTICAL_CENTER);
    // 内容水平居中
    format_set_align(columnDetailformat, LXW_ALIGN_CENTER);
    // 遍历数据源，将表格数据写入到某行某列中
    for (int i = 0; i < self.accountDataArray.count; i++) {
        NSDictionary *model = self.accountDataArray[i];
        
        NSDate *date = [NSDate date];
        NSDateFormatter *formatter = [[NSDateFormatter alloc] init];
        [formatter setDateFormat:@"yyyy-MM-dd HH-mm-ss"];//样式
        NSString *time_now = [formatter stringFromDate:date];
        
        worksheet_set_row(worksheet, 3+i, 30, columnTitleformat);
        worksheet_write_string(worksheet, 3+i, 0, [[NSString stringWithFormat:@"%d",i+1] UTF8String], columnDetailformat);
        worksheet_write_string(worksheet, 3+i, 1, [model[@"companyName"] UTF8String], columnDetailformat);
        worksheet_write_string(worksheet, 3+i, 2, [model[@"accountNo"] UTF8String], columnDetailformat);
        worksheet_write_string(worksheet, 3+i, 3, [time_now UTF8String], columnDetailformat);
        worksheet_write_string(worksheet, 3+i, 4, [model[@"orderNo"] UTF8String], columnDetailformat);
        worksheet_write_string(worksheet, 3+i, 5, [model[@"price"] UTF8String], columnDetailformat);
        worksheet_write_string(worksheet, 3+i, 6, [model[@"allInPayOrderNo"] UTF8String], columnDetailformat);
        worksheet_write_string(worksheet, 3+i, 7, [model[@"discountMoney"] UTF8String], formatNum);
        worksheet_write_number(worksheet, 3+i, 8, [model[@"payment"] doubleValue], formatNum);
    }
    
    NSString *sumStr = [NSString stringWithFormat:@"=SUM(I4:I%lu)",3+self.accountDataArray.count];
    // 采用数学公式计算
    worksheet_write_formula(worksheet, rowA, 8, [sumStr cStringUsingEncoding:NSUTF8StringEncoding], formatNum);
    
    //保存
    workbook_close(workbook);
    
    // 用于webview展示
    dispatch_after(dispatch_time(DISPATCH_TIME_NOW, (int64_t)(1 * NSEC_PER_SEC)), dispatch_get_main_queue(), ^{
        
        NSString *documentPath = [NSSearchPathForDirectoriesInDomains(NSDocumentDirectory, NSUserDomainMask, YES) objectAtIndex:0];
        NSString *file1Name = [NSString stringWithFormat:@"%ld年%ld月%ld日-%ld年%ld月%ld日对账单.xlsx",self.startYear,self.startMonth,self.startDay,self.endYear,self.endMonth,self.endDay];
        NSString *filePath = [documentPath stringByAppendingPathComponent:file1Name];
        NSURL *url = [NSURL fileURLWithPath:filePath]; // 注意：使用[NSURL URLWithString:filePath]无效
        NSURLRequest *urlRequest = [NSURLRequest requestWithURL:url];
        [self.testWebView loadRequest:urlRequest];
    });
    
}

// 导出按钮点击
-(void)exportAccountClick:(UIButton *)btn{
    // 获取cache路径
    NSString *documentDirectory = [NSSearchPathForDirectoriesInDomains(NSDocumentDirectory, NSUserDomainMask, YES) objectAtIndex:0];
    NSString *xlsxName = [NSString stringWithFormat:@"%ld年%ld月%ld日-%ld年%ld月%ld日对账单.xlsx",self.startYear,self.startMonth,self.startDay,self.endYear,self.endMonth,self.endDay];
    // 拼接 cache路径 + 名为xlsxName的文件。注意：注意filename一定要和当初创建时的文件名保持一致，否则读取不到该文件
    NSString *filename = [documentDirectory stringByAppendingPathComponent:xlsxName];
    // 初始化并读取路径为filename的文件
    self.document = [UIDocumentInteractionController interactionControllerWithURL:[NSURL fileURLWithPath:filename]];
    self.document.delegate = self;
    // 以present的方式展示该文件所支持的第三方App列表
    [self.document presentOpenInMenuFromRect:self.view.bounds inView:self.view animated:YES];
}


#pragma mark - documentDelegate

-(UIViewController*)documentInteractionControllerViewControllerForPreview:(UIDocumentInteractionController*)controller{
    return self;
}

-(UIView*)documentInteractionControllerViewForPreview:(UIDocumentInteractionController*)controller {
    return self.view;
}

- (CGRect)documentInteractionControllerRectForPreview:(UIDocumentInteractionController*)controller {
    return self.view.frame;
}

//点击预览窗口的“Done”(完成)按钮时调用
- (void)documentInteractionControllerDidEndPreview:(UIDocumentInteractionController*)controller {
    
}

// 文件分享面板弹出的时候调用
- (void)documentInteractionControllerWillPresentOpenInMenu:(UIDocumentInteractionController*)controller{
    
    NSLog(@"WillPresentOpenInMenu");
    
}

// 当选择一个文件分享App的时候调用
- (void)documentInteractionController:(UIDocumentInteractionController*)controller willBeginSendingToApplication:(nullable NSString*)application{
    
    NSLog(@"begin send : %@", application);
    
}

// 弹框消失的时候走的方法
-(void)documentInteractionControllerDidDismissOpenInMenu:(UIDocumentInteractionController*)controller{
    
    NSLog(@"dissMiss");
    
}

#pragma mark - webView
- (WKWebView *)testWebView{
    if (_testWebView) return _testWebView;
    _testWebView = [[WKWebView alloc]initWithFrame:CGRectMake(10, NAV_HEIGHT, APP_SCREEN_WIDTH-20, APP_SCREEN_HEIGHT-NAV_HEIGHT-BOTTOM_SAFE_AREA-100)];
    _testWebView.navigationDelegate = self;
    _testWebView.UIDelegate = self;
    _testWebView.backgroundColor = [UIColor whiteColor];
    [self.view addSubview:_testWebView];
    return _testWebView;
}
#pragma mark - WKWebView Delegate
// 页面开始加载时调用
- (void)webView:(WKWebView *)webView didStartProvisionalNavigation:(WKNavigation *)navigation{
    NSLog(@"页面开始加载时调用");
}

// 当内容开始返回时调用
- (void)webView:(WKWebView *)webView didCommitNavigation:(WKNavigation *)navigation{
    NSLog(@"当内容开始返回时调用");
}

// 页面加载完成之后调用
- (void)webView:(WKWebView *)webView didFinishNavigation:(WKNavigation *)navigation{
    NSLog(@"页面加载完成之后调用");
}

// 页面加载失败时调用
- (void)webView:(WKWebView *)webView didFailProvisionalNavigation:(WKNavigation *)navigation{
    NSLog(@"页面加载失败时调用");
}

// 接收到服务器跳转请求之后调用
- (void)webView:(WKWebView *)webView didReceiveServerRedirectForProvisionalNavigation:(WKNavigation *)navigation{
    NSLog(@"接收到服务器跳转请求之后调用");
}

@end
