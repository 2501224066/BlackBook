<?php
ini_set("display_errors", "Off");

if( (!isset($_GET['code'])) || ($_GET['code'] != (date('Ymd')-996))) {
    echo "<div style='width: 100%; text-align: center;position: fixed; top: 40%;'>";
    echo "<h1>BLACK BOOK</h1>";
    echo "<a href='index.php?code=输入暗号'>Tomorrow, where will we </a>";
    echo "</div>";
    exit;
}

/*读取excel文件，并进行相应处理*/
$fileName = "BlackBook.xlsx";
if (!file_exists($fileName))
    exit("文件".$fileName."不存在");

$startTime = time(); //返回当前时间的Unix 时间戳

require_once("./PHPExcel/Classes/PHPExcel.php");

$reader = new PHPExcel_Reader_Excel2007();

$PHPExcel = $reader->load($fileName, 'utf-8'); // 载入excel文件
$sheet = $PHPExcel->getSheet(0); // 读取第一個工作表
$highestRow = $sheet->getHighestRow(); // 取得总行数
$highestColumm = $sheet->getHighestColumn(); // 取得总列数

$data = array();
for ($rowIndex = 1; $rowIndex <= $highestRow; $rowIndex++) {        //循环读取每个单元格的内容。注意行从1开始，列从A开始
    for ($colIndex = 'A'; $colIndex <= $highestColumm; $colIndex++) {
        $addr = $colIndex . $rowIndex;
        $cell = $sheet->getCell($addr)->getValue();
        if ($cell instanceof PHPExcel_RichText) { //富文本转换字符串
            $cell = $cell->__toString();
        }
        $data[$rowIndex][$colIndex] = $cell;
    }
}

?>

<!DOCTYPE html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- 上述3个meta标签*必须*放在最前面，任何其他内容都*必须*跟随其后！ -->
    <title>BlackBook</title>

    <!-- Bootstrap -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@3.3.7/dist/css/bootstrap.min.css" rel="stylesheet">

    <!-- HTML5 shim 和 Respond.js 是为了让 IE8 支持 HTML5 元素和媒体查询（media queries）功能 -->
    <!-- 警告：通过 file:// 协议（就是直接将 html 页面拖拽到浏览器中）访问页面时 Respond.js 不起作用 -->
    <!--[if lt IE 9]>
      <script src="https://cdn.jsdelivr.net/npm/html5shiv@3.7.3/dist/html5shiv.min.js"></script>
      <script src="https://cdn.jsdelivr.net/npm/respond.js@1.4.2/dest/respond.min.js"></script>
    <![endif]-->
  </head>
  <body>
    <br/>
    <div class="container">
        <table class="table table-hover table-bordered">
            <tbody>
                <?php
                foreach ($data as $k => $v) {
                    if ($k != 1) {
                        echo "<tr>";
                        echo "<td>" . $v['A'] . "</td>";
                        echo "<td><a href='https://map.baidu.com/search/" . $v['B'] . "/@12734330.915513515,3546225.025,15.21z?querytype=con&wd=%E5%85%89%E8%B0%B7%E5%B9%BF%E5%9C%BA&c=218&provider=pc-aladin&pn=0&device_ratio=1&da_src=shareurl'>" . $v['B'] . "</a></td>";
                        echo "<td>" . $v['C'] . "</td>";
                        echo "</tr>";
                    } else {
                        echo "<tr>";
                        echo "<th>" . $v['A'] . "</th>";
                        echo "<th>" . $v['B'] . "</th>";
                        echo "<th>" . $v['C'] . "</th>";
                        echo "</tr>";
                    }
                }
                ?>
            </tbody>
        </table>
    </div>

    <!-- jQuery (Bootstrap 的所有 JavaScript 插件都依赖 jQuery，所以必须放在前边) -->
    <script src="https://cdn.jsdelivr.net/npm/jquery@1.12.4/dist/jquery.min.js"></script>
    <!-- 加载 Bootstrap 的所有 JavaScript 插件。你也可以根据需要只加载单个插件。 -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@3.3.7/dist/js/bootstrap.min.js"></script>
  </body>
</html>