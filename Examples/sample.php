
<?php
include '../Classes/PHPTemplate.php';
include '../vendor/autoload.php';


class Foo
{
	private $_var1;
	private $_var2;
	private $_var3;

	public function setVar1($param) {
		$this->_var1 = $param;
	}
	public function setVar2($param) {
		$this->_var2 = $param;
	}
	public function setVar3($param) {
		$this->_var3 = $param;
	}

	public function getVar1() {
		return $this->_var1;
	}
	public function getVar2() {
		return $this->_var2;
	}
	public function getVar3() {
		return $this->_var3;
	}

}


// 新規作成の場合
$obj = new PHPTemplate();

$data = array();
$data['test1'] = "aiueo";
$data['test2'] = "12345";
$data['test3'] = "あいうえお";
$data['test4'] = "アイウエオ";
$data['test5'] = "亜委鵜得尾";
$data['test6'] = "!\"#$";
$data['test7'] = "123/いあうおえ";
$data['test8'] = 54321;
$date = new DateTime("2014-11-22 14:41:00");
$data['date1'] = $date;
$testArray = array();
$a = new Foo();
$a->setVar1("<test>フォント適用</test>ああああ");
$a->setVar2("listdata1");
$a->setVar3("jiba.jpeg");
$testArray[] = $a;
$b = new Foo();
$b->setVar1("<test>フォント適用</test>いいいい");
$b->setVar2("listdata2");
$b->setVar3("jiba.jpeg");
$testArray[] =$b;
$c = new Foo();
$c->setVar1("<test>フォント適用</test>うううう");
$c->setVar2("listdata3");
$c->setVar3("jiba.jpeg");
$testArray[] = $c;
$d = new Foo();
$d->setVar1("<test>フォント適用</test>uuuuuuuuu");
$d->setVar2("listdata4");
$d->setVar3("jiba.jpeg");
$testArray[] = $d;
$data['list1'] = $testArray;

$testArray2 = array();
$a2 = new Foo();
$a2->setVar1("abcde");
$a2->setVar2("12345");
$testArray2[] = $a2;
$b2 = new Foo();
$b2->setVar1("fghij");
$b2->setVar2("6789");
$testArray2[] =$b2;
$data['list2'] = $testArray2;


$foo = new Foo();
$foo->setVar1("ばー１");
$foo->setVar2("ばー２");
$data['foo'] = $foo;
$data['pict1'] = "jiba.jpeg";


$data['yy'] = "15";
$data['mm'] = "02";

$array =array(
			array(
				"data1" => "111",
				"data2" => "222",
				"data3" => 123,
			),
			array(
				"data1" => "333",
				"data2" => "444",
				"data3" => 456,
			)
		);
$data['result2']=$array;

$array = array();
for($i=0;$i<3;$i++){
	$dataArray = array();
	$dataArray['data1'] = "data1".$i;
	$dataArray['data2'] = "data2".$i;
	$dataArray['data3'] = $i;
	$array[] = $dataArray;
}
$data['result']=$array;

$data['flag']=true;

//print_r ($obj->getSheetName("test.xls", PHPTemplate_Util::EXCEL_TYPE_95));
// Excel95の場合
//$obj->writeExcel("test.xls", $data, "out.xls", PHPTemplate_Util::EXCEL_TYPE_95);

// Excel2007の場合
$obj->writeExcel("test2007.xlsx", $data, "out2007.xlsx", PHPTemplate_Util::EXCEL_TYPE_2007);

//$obj->htmlTest("Book2.htm", "Book2.xlsx");

//シート追加
//$obj->addSheetByName("test2007_sheet.xlsx", "baseSheet", "addSheet");
/*$replaceValues = array();
$replace = new PHPTemplate_ReplaceValue("fukatest", "replace");
$replaceValues[] = $replace;
$replace = new PHPTemplate_ReplaceValue("test2", "test2_rep");
$replaceValues[] = $replace;*/
//print_r($replaceValues);
//$obj->addSheet("test2007_sheet.xlsx", 0, "addSheetTest", NULL, NULL, $replaceValues);
//$obj->addSheetByName("test2007_sheet.xlsx", "Sheet2", "addSheetTest", NULL, NULL, $replaceValues);
//シート削除
//$obj->removeSheet("test2007_sheet.xlsx", 1);


// phpinfo();


//$obj->test("test2007_sheet.xlsx", "fuka_test.xlsx");
echo 'end';
