### 生成Excel

#### 项目说明
```
本库用于通过数组的形式生成 Excel， 并可将Excel发送到浏览器，写入本地文件等。
本库是基于 https://github.com/rogeriopradoj/php-excel.git 进行的改进。
修改的功能如下：
1. 代码中添加了中文注释
2. 文件名支持中文
3. 去掉了一些兼容性的代码
```

#### 使用示例
-----
添加依赖 ``jingwu/excel`` 到项目的 ``composer.json`` 文件:
```json
    {
        "require": {
            "jingwu/excel": "0.1.2"
        }
    }
```

```
require "vendor/autoload.php";

use Jingwu\Excel\Excel_XML;

$data = array(
    0 => array('ID', '用户名', '邮箱'),
    array(1, '张三', '100000@qq.com'),
    array(2, '李四', '100001@qq.com')
);

$xls = new Excel_XML;
$xls->addWorksheet('Names', $data);
$xls->sendWorkbook('test.xml');
//$xls->writeWorkbook('test.xml');
```

