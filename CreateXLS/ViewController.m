//
//  ViewController.m
//  CreateXLS
//
//  Created by Mr.Liu on 16/12/8.
//  Copyright © 2016年 Personal. All rights reserved.
//

#import "ViewController.h"

@interface ViewController ()

@end

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
    // Do any additional setup after loading the view, typically from a nib.
    
    UIButton *createButton = [UIButton buttonWithType:UIButtonTypeSystem];
    createButton.frame = CGRectMake(130, 300, 100, 20);
    [createButton setTitle:@"生成XLS文件" forState:UIControlStateNormal];
    [createButton addTarget:self action:@selector(createXLSFile) forControlEvents:UIControlEventTouchUpInside];
    
    [self.view addSubview:createButton];
}

- (void)createXLSFile {
    
    // 创建存放XLS文件数据的数组
    NSMutableArray  *xlsDataMuArr = [[NSMutableArray alloc] init];
    
    // 第一行内容
    [xlsDataMuArr addObject:@"Time"];
    [xlsDataMuArr addObject:@"Address"];
    [xlsDataMuArr addObject:@"Person"];
    [xlsDataMuArr addObject:@"Origin"];
    [xlsDataMuArr addObject:@"Process"];
    [xlsDataMuArr addObject:@"Result"];
    
    // 100行数据
    for (int i = 0; i < 100; i ++) {
        
        [xlsDataMuArr addObject:@"2016-12-06 17:18:40"];
        [xlsDataMuArr addObject:@"GuangZhou"];
        [xlsDataMuArr addObject:@"Mr.Liu"];
        [xlsDataMuArr addObject:@"Buy"];
        [xlsDataMuArr addObject:@"TaoBao"];
        [xlsDataMuArr addObject:@"Debt"];
    }
    
    // 数组转换为字符串，连接符：@"\t"（功能同键盘上的tab键）
    NSString *fileContent = [xlsDataMuArr componentsJoinedByString:@"\t"];
    
    // 字符串转换为可变字符串，方便改变其中的字符
    NSMutableString *muStr = [NSMutableString stringWithString:fileContent];
    
    // 新建一个可变数组，存储每行最后一个\t的下标
    NSMutableArray *rangeMuArr = [[NSMutableArray alloc] init];
    
    for (int i = 0; i < muStr.length; i ++) {
        
        NSRange range = [muStr rangeOfString:@"\t" options:NSLiteralSearch range:NSMakeRange(i, 1)];
        
        if (range.length == 1) {
            
            [rangeMuArr addObject:@(range.location)];
        }
    }
    
    // 每行最后一个\t替换为\n
    for (NSUInteger i = 0; i < rangeMuArr.count; i ++) {
#warning  下面的6是列数，根据需求修改
        if ( i > 0 && (i%6 == 0) ) {
            
            [muStr replaceCharactersInRange:NSMakeRange([[rangeMuArr objectAtIndex:i-1] integerValue], 1) withString:@"\n"];
        }
    }
    
    // 文件管理器
    NSFileManager *fileManager = [[NSFileManager alloc]init];
    
    //使用UTF16才能显示汉字；如果显示为#######是因为格子宽度不够，拉开即可
    NSData *fileData = [muStr dataUsingEncoding:NSUTF16StringEncoding];
    
    // 文件路径
    NSString *path = NSHomeDirectory();
    NSString *filePath = [path stringByAppendingPathComponent:@"export.xls"];
    NSLog(@"文件路径：\n%@",filePath);
    // 生成xls文件
    [fileManager createFileAtPath:filePath contents:fileData attributes:nil];
}

- (void)didReceiveMemoryWarning {
    [super didReceiveMemoryWarning];
    // Dispose of any resources that can be recreated.
}

@end
