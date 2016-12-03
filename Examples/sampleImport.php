
<?php
include '../Classes/PHPTemplate.php';
include '../vendor/autoload.php';


// Excelデータ読みこみサンプル
// 設定ファイル(xml)に従い、エクセルファイルを読み込む(readExcel)
// read後、フィルター条件を渡し、フィルタリングした結果をArrayにて返す(filter)
$obj = new PHPTemplate();
// Excel95の場合
$obj->readExcel("importTestData.xls", "importConfig.xml", PHPTemplate_Util::EXCEL_TYPE_95);
// Excel2007の場合
//$obj->readExcel("importTestData2007.xlsx", "importConfig.xml", PHPTemplate_Util::EXCEL_TYPE_2007);

$filter = new PHPTemplate_Filter();

$columnFilter1 = new PHPTemplate_FilterColumn('column1', PHPTemplate_Util::FILTER_TYPE_TEXT , 'ああああ１');
$filter->setColumnFilter(1, $columnFilter1);/*シートインデックスとフィルターをセット*/

$columnFilter2 = new PHPTemplate_FilterColumn('column3', PHPTemplate_Util::FILTER_TYPE_CELL_COLOR , 'FFFF00');
$filter->setColumnFilter(1, $columnFilter2);

$columnFilter3 = new PHPTemplate_FilterColumn('column4', PHPTemplate_Util::FILTER_TYPE_FONT_COLOR , 'FF0000');
$columnFilter4 = new PHPTemplate_FilterColumn('column1', PHPTemplate_Util::FILTER_TYPE_TEXT , 'ああああ２');
$arrayFilter = Array(); // Arrayで渡すとそれぞれの条件をANDで判定する
$arrayFilter[] = $columnFilter3;
$arrayFilter[] = $columnFilter4;
$filter->setColumnFilter(1, $arrayFilter);

$columnFilter5 = new PHPTemplate_FilterColumn('カラム４', PHPTemplate_Util::FILTER_TYPE_FONT_COLOR , 'FF0000');
$filter->setColumnFilter(2, $columnFilter5);

// 条件設定されていないシートは全件返す
$filterData[] = $obj->filter($filter);
print_r($filterData);

echo 'end';
