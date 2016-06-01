<?php 
header('Content-Type: text/html; charset=UTF-8');

ini_set('memory_limit', '1024M');
set_time_limit(0);

define('ModifyColNum', 'B');

//首先导入PHPExcel
require_once 'PHPExcel.php';
require_once 'PHPExcel/IOFactory.php';
require_once 'RollingCurl.php';

//设置当前需要处理的文档
$filePath = "2.xlsx";


$citys = array();
$htmls = array();

//建立reader对象
$PHPReader = new PHPExcel_Reader_Excel2007();
if(!$PHPReader->canRead($filePath)){
    $PHPReader = new PHPExcel_Reader_Excel5();
    if(!$PHPReader->canRead($filePath)){
        echo 'no Excel';
        return ;
    }
}

//建立excel对象，此时你即可以通过excel对象读取文件，也可以通过它写入文件
$PHPExcel = $PHPReader->load($filePath);
/**读取excel文件中的第一个工作表*/
$currentSheet = $PHPExcel->getSheet(0);
/**取得最大的列号*/
$allColumn = $currentSheet->getHighestColumn();
/**取得一共有多少行*/
$allRow = $currentSheet->getHighestRow();

$edata = array();

$rc = new RollingCurl("GetHtml");
$rc->window_size = 20;

//循环读取每个单元格的内容。注意行从1开始，列从A开始
for($colIndex='A';$colIndex<=$allColumn;$colIndex++){
    for($rowIndex=1;$rowIndex<=$allRow;$rowIndex++){
        $addr = $colIndex.$rowIndex;
        $cell = $currentSheet->getCell($addr)->getValue();
        if($cell instanceof PHPExcel_RichText)     //富文本转换字符串
            $cell = $cell->__toString();
        if ($colIndex == 'C') {
            $search = "http://www.so.com/s?ie=utf-8&shb=1&src=home_so.com&q=".urlencode($cell);
            $extra = [
                'company' => $cell,
                'rowIndex' => $rowIndex,
                'alterIndex' => ModifyColNum.$rowIndex,
            ];
            $request = new RollingCurlRequest($search, 'GET', null, null, null, $extra);
            $rc->add($request);
            // $company[$rowIndex] = $cell;
        }
        $edata[$rowIndex][$colIndex] = $cell;
    }
}


try {
    $rc->execute();
} catch (Exception $e) {
    throw new Exception($e->getMessage());
}

//批量将获取到的文本信息分词后扑捉第一个城市信息
GetCitys($htmls);

//将数据放回到数组中
foreach ($edata as $k=>$v) {
    $key = ModifyColNum.$k;
    $edata[$k][ModifyColNum] = @$citys[$key];
}

$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setTitle("export")->setDescription("none");  
$objPHPExcel->setActiveSheetIndex(0); 

// Field names in the first row  
$col = 0;
// Fetching the table data  
$row = 1;

foreach($edata as $k=>$v)  
{  
    $col = 0;  
    foreach ($v as $i)
    {
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, $i);  
        $col++;
    }
    $row++; 
}  

$objPHPExcel->setActiveSheetIndex(0);
$objWriter = IOFactory::createWriter($objPHPExcel, 'Excel5');  

// Sending headers to force the user to download the file  
header('Content-Type: application/vnd.ms-excel');  
header('Content-Disposition: attachment;filename="'.time().'.xls"');  
header('Cache-Control: max-age=0');  
$objWriter->save('php://output');


exit('success');


/**
 * 批量获取curl后的文本信息
 * @param [type] $response [description]
 * @param [type] $info     [description]
 * @param [type] $request  [description]
 */
function GetHtml($response, $info, $request)
{
    global $htmls;
    $htmls[$request->extra['alterIndex']] = $response;
}

function GetCitys($htmls)
{
    $rc = new RollingCurl("GetCity");
    $rc->window_size = 20;

    foreach ($htmls as $k=>$v) {
        /*此处的匹配条件需要根据上一步的搜索工具的改变而改变，目前使用的是so.com的搜索结果*/
        preg_match("/<div id=\"main\">([\s\S]*?)<!-- END #main -->/", $v, $m);
        $s = strip_tags(@$m[1]);
        /*POST Method*/        
        $fenci = "http://api.pullword.com/post.php";
        $extra = [
            'alterIndex' => $k,
        ];
        $pdata = array(
            'source' => $s,
            'param1' => 0,
            'param2' => 0,
        );
        $request = new RollingCurlRequest($fenci, 'POST', $pdata, null, null, $extra);

        /*GET Method*/
        // $fenci = "http://api.pullword.com/get.php?source=".$s."&param1=0&param2=0";
        // $extra = [
        //     'alterIndex' => $k,
        // ];
        // $request = new RollingCurlRequest($fenci, 'GET', null, null, null, $extra);

        $rc->add($request);
    }
    try {
        $rc->execute();
    } catch (Exception $e) {
        throw new Exception($e->getMessage());
    }
}


function GetCity($response, $info, $request)
{
    global $citys;
    $city = t($response);
    $citys[$request->extra['alterIndex']] = $city;
}


function t($s)
{
    $t = <<<EOF
北京
天津
河北
山西
内蒙古
辽宁
吉林
黑龙江
上海
江苏
浙江
安徽
福建
江西
山东
河南
湖北
湖南
广东
广西
海南
重庆
四川
贵州
云南
西藏
陕西
甘肃
青海
宁夏
新疆
香港
澳门
台湾
和平区
石家庄
唐山
秦皇岛
邯郸
邢台
保定
张家口
承德
沧州
廊坊
衡水
太原
大同
阳泉
长治
晋城
朔州
晋中
运城
忻州
临汾
吕梁
呼和浩特
包头
乌海
赤峰
通辽
鄂尔多斯
呼伦贝尔
巴彦淖尔
乌兰察布
兴安盟
锡林郭勒盟
阿拉善盟
沈阳
大连
鞍山
抚顺
本溪
丹东
锦州
营口
阜新
辽阳
盘锦
铁岭
朝阳
葫芦岛
长春
吉林
四平
辽源
通化
白山
松原
白城
延边朝鲜族自治州
哈尔滨
齐齐哈尔
鸡西
鹤岗
双鸭山
大庆
伊春
佳木斯
七台河
牡丹江
黑河
绥化
大兴安岭地区
南京
无锡
徐州
常州
苏州
南通
连云港
淮安
盐城
扬州
镇江
泰州
宿迁
杭州
宁波
温州
嘉兴
湖州
绍兴
金华
衢州
舟山
台州
丽水
合肥
芜湖
蚌埠
淮南
马鞍山
淮北
铜陵
安庆
黄山
滁州
阜阳
宿州
巢湖
六安
亳州
池州
宣城
福州
厦门
莆田
三明
泉州
漳州
南平
龙岩
宁德
南昌
景德镇
萍乡
九江
新余
鹰潭
赣州
吉安
宜春
抚州
上饶
济南
青岛
淄博
枣庄
东营
烟台
潍坊
济宁
泰安
威海
日照
莱芜
临沂
德州
聊城
滨州
荷泽
郑州
开封
洛阳
平顶山
安阳
鹤壁
新乡
焦作
濮阳
许昌
漯河
三门峡
南阳
商丘
信阳
周口
驻马店
武汉
黄石
十堰
宜昌
襄樊
鄂州
荆门
孝感
荆州
黄冈
咸宁
随州
恩施土家族苗族自治州
神农架
长沙
株洲
湘潭
衡阳
邵阳
岳阳
常德
张家界
益阳
郴州
永州
怀化
娄底
湘西土家族苗族自治州
广州
韶关
深圳
珠海
汕头
佛山
江门
湛江
茂名
肇庆
惠州
梅州
汕尾
河源
阳江
清远
东莞
中山
潮州
揭阳
云浮
南宁
柳州
桂林
梧州
北海
防城港
钦州
贵港
玉林
百色
贺州
河池
来宾
崇左
海口
三亚
黄浦区
成都
自贡
攀枝花
泸州
德阳
绵阳
广元
遂宁
内江
乐山
南充
眉山
宜宾
广安
达州
雅安
巴中
资阳
阿坝藏族羌族自治州
甘孜藏族自治州
凉山彝族自治州
贵阳
六盘水
遵义
安顺
铜仁地区
黔西南布依族苗族自治州
毕节地区
黔东南苗族侗族自治州
黔南布依族苗族自治州
昆明
曲靖
玉溪
保山
昭通
丽江
思茅
临沧
楚雄彝族自治州
红河哈尼族彝族自治州
文山壮族苗族自治州
西双版纳傣族自治州
大理白族自治州
德宏傣族景颇族自治州
怒江傈僳族自治州
迪庆藏族自治州
拉萨
昌都地区
山南地区
日喀则地区
那曲地区
阿里地区
林芝地区
西安
铜川
宝鸡
咸阳
渭南
延安
汉中
榆林
安康
商洛
兰州
嘉峪关
金昌
白银
天水
武威
张掖
平凉
酒泉
庆阳
定西
陇南
临夏回族自治州
甘南藏族自治州
西宁
海东地区
海北藏族自治州
黄南藏族自治州
海南藏族自治州
果洛藏族自治州
玉树藏族自治州
海西蒙古族藏族自治州
银川
石嘴山
吴忠
固原
中卫
乌鲁木齐
克拉玛依
吐鲁番地区
哈密地区
昌吉回族自治州
博尔塔拉蒙古自治州
巴音郭楞蒙古自治州
阿克苏地区
克孜勒苏柯尔克孜自治州
喀什地区
和田地区
伊犁哈萨克自治州
塔城地区
阿勒泰地区
石河子
阿拉尔
图木舒克
五家渠
香港特别行政区
澳门特别行政区
台湾
东城区
西城区
崇文区
朝阳区
丰台区
石景山区
海淀区
门头沟区
房山区
通州区
顺义区
昌平区
大兴区
怀柔区
平谷区
密云县
延庆县
卢湾区
徐汇区
长宁区
静安区
普陀区
闸北区
虹口区
杨浦区
闵行区
宝山区
嘉定区
浦东新区
金山区
松江区
青浦区
南汇区
奉贤区
崇明县
河东区
河西区
南开区
河北区
红桥区
塘沽区
汉沽区
大港区
东丽区
西青区
津南区
北辰区
武清区
宝坻区
宁河县
静海县
蓟县
忻府区
定襄县
五台县
代县
繁峙县
宁武县
静乐县
神池县
五寨县
岢岚县
河曲县
保德县
偏关县
原平
EOF
;


$AllCitysArr = explode(PHP_EOL, $t);
$ContentArr = explode(PHP_EOL, $s);


foreach ($ContentArr as $k=>$v) {

    foreach ($AllCitysArr as $i=>$j) {

        if (preg_match("/$j/", $v)) {
            return $j;
        }

    }
}
return '';
}


?>