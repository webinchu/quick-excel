# quick-excel

PHP快速生成excel文档

- 示例代码

```
//表头
$excelHead = ['id', '名称', '时间'];
$data = [];  //多维数组或者model类
for ($i = 0; $i <= 5; $i++) {
    $data[$i] = [
        $i + 1,
        '名称' . ($i + 1),
        date('Y-m-d H:i:s', time())
    ];
}
$fileName = '测试';
$excel = new QuickExcel();
//自动下载
$excel::excelOut($excelHead, $data, $fileName);
```

- 如图:

![img.png](img.png)
