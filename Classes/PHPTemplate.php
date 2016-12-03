<?php
/**
 * PHPTemplate
 *
 * @category   PHPTemplate
 * @package    PHPTemplate
 * @copyright
 */

if (!defined('PHPTEMPLATEL_ROOT')) {
    define('PHPTEMPLATE_ROOT', dirname(__FILE__) . '/');
    require(PHPTEMPLATE_ROOT . 'PHPTemplate/Autoloader.php');
}

class PHPTemplate
{

	/**
	 * 読み込み済みデータ
	 * @var	Array
	 */
	private $_roadData = Array();

	/**
	 * シート数取得
	 *
	 * @return int
	 */
	public function getSheetCount(){
		return count($this->_roadData);
	}

	/**
	 * シート毎のデータ取得
	 * @return Array
	 */
	public function getSheetData($idx){
		return $this->_roadData[$idx];
	}

	/**
	 *	テンプレート用ワークブックと埋め込み用データを受け取り、ファイル出力
	 *
	 *	@param	string	  			$filepath
	 *	@param	array  				$data
	 *	@param	string			 	$outpath
	 *  @param	string				$excelType(PHPTemplate_Util::EXCEL_TYPE_95,EXCEL_TYPE_2007)
	 */
	public function writeExcel($filepath, $data, $outpath, $excelType=PHPTemplate_Util::EXCEL_TYPE_2007){
		$reader = PHPExcel_IOFactory::createReader($excelType);
		$objPHPExcel = $reader->load($filepath);

		$sheets = $this->getWorksheets($objPHPExcel);

		if (empty($sheets)) {
			return NULL;
		}

		foreach ($sheets as $sheet) {

			$objSheet = new PHPTemplate_Worksheet($sheet);
			$objSheet->analyzeSheet($data);

		}
		// Excel 95 形式
		$writer = PHPExcel_IOFactory::createWriter($objPHPExcel, $excelType);
		$writer->save($outpath);

		return $outpath;
	}

	/**
	 *	ワークブックからシートを取得
	 *
	 *	@param	PHPExcel  			$objPHPExcel
	 *	@param	int  				$sheetIndex
	 *	@return	array
	 */
	private function getWorksheets($objPHPExcel, $sheetIndex = NULL)
	{
		$sheets = array();
		if (is_null($sheetIndex)) {
			$sheets = $objPHPExcel->getAllSheets();
		} elseif (is_array($sheetIndex)) {
			foreach($sheetIndex as $idx) {
				$sheets[$idx] = $objPHPExcel->getSheet($idx);
			}
		} elseif (is_int($sheetIndex)) {
			$sheets[$sheetIndex] = $objPHPExcel->getSheet($sheetIndex);
		}
		return $sheets;
	}

	/**
	 *	ワークブックのシート名を取得
	 *
	 *	@param	PHPExcel  			$filepath
	 *	@return	array
	 */
	public function getSheetName($filepath, $sheetIndex=NULL, $excelType=PHPTemplate_Util::EXCEL_TYPE_2007)
	{
		$reader = PHPExcel_IOFactory::createReader($excelType);
		$objPHPExcel = $reader->load($filepath);

		$sheetsNames = array();
		if (is_null($sheetIndex)) {
			$sheetObj = array();
			$sheetObj = $objPHPExcel->getAllSheets();
			foreach ($sheetObj as $s){
				$sheetsNames[] = $s->getTitle();
			}
		} elseif (is_array($sheetIndex)) {
			foreach($sheetIndex as $idx) {
				$sheetsNames[$idx] = $objPHPExcel->getSheet($idx)->getTitle();
			}
		} elseif (is_int($sheetIndex)) {
			$sheetsNames[$sheetIndex] = $objPHPExcel->getSheet($sheetIndex)->getTitle();
		}
		return $sheetsNames;
	}



	/**
	 *	インポート用ファイルを受け取り設定に従ってデータを返す
	 *
	 *	@param	string	  		$importFilePath
	 *	@param	string  		$configFilePath
	 *  @param	string			$excelType(PHPTemplate_Util::EXCEL_TYPE_95,EXCEL_TYPE_2007)
	 */
	public function readExcel($importFilePath, $configFilePath, $excelType=PHPTemplate_Util::EXCEL_TYPE_2007)
	{
		// 設定ファイル(XML)を読み込む
		$config = new PHPTemplate_Config();
		$config->readXML($configFilePath);

		//エクセル読み込み
		$reader = PHPExcel_IOFactory::createReader($excelType);
		$objPHPExcel = $reader->load($importFilePath);

		$sheets = $this->getWorksheets($objPHPExcel);

		if (empty($sheets)) {
			return NULL;
		}
		$sheetIndex = 1;
		foreach ($sheets as $sheet) {
			$conf = $config->getSheetConfig($sheetIndex);
			if($conf != NULL){
				$sheetData = new PHPTemplate_ImportData($sheet, $conf, $sheetIndex);
				$sheetData->analyzeSheet();
				$this->_roadData[$sheet->getTitle()]=$sheetData;
			}
			$sheetIndex++;
		}
		return $this->_roadData;
	}

	/**
	 *	インポートデータをフィルタリング
	 *  (値・文字色・セル色)含む行を返す
	 *	@param		  PHPTemplate_Filter		$filter
	 */
	public function filter($filter)
	{
		$outSheetArray=Array();
		foreach ($this->_roadData as $sheetData) {

			$index = $sheetData->getSheetIndex();
			$filterData = $filter->getSheetFilter($index);
			$outSheetArray[]=$sheetData->getFilterData($filterData);

		}
		return $outSheetArray;
	}

	/**
	 *	ワークブックのシートを削除
	 *
	 *	@param	$filepath
	 *  @param  $sheetIndex
	 *  @param  $excelType
	 *	@return
	 */
	public function removeSheet($filepath, $sheetIndex, $outpath=NULL, $excelType=PHPTemplate_Util::EXCEL_TYPE_2007)
	{
		$reader = PHPExcel_IOFactory::createReader($excelType);
		$objPHPExcel = $reader->load($filepath);
		if (is_int($sheetIndex)) {
			if($objPHPExcel->getSheetCount() > $sheetIndex){
				$objPHPExcel->removeSheetByIndex($sheetIndex);
				if($outpath==NULL){
					$outpath = $filepath;
				}
				$writer = PHPExcel_IOFactory::createWriter($objPHPExcel, $excelType);
				$writer->save($outpath);
				return $outpath;
			}
		}
		return NULL;

	}

	/**
	 *	ワークブックのシートを削除
	 *
	 *	@param	$filepath
	 *  @param  $sheetName
	 *  @param  $excelType
	 *	@return
	 */
	public function removeSheetByName($filepath, $sheetName, $outpath=NULL, $excelType=PHPTemplate_Util::EXCEL_TYPE_2007)
	{
		$reader = PHPExcel_IOFactory::createReader($excelType);
		$objPHPExcel = $reader->load($filepath);

		$sheet = $objPHPExcel->getSheetByName($sheetName);
		if($sheet != NULL){
			$objPHPExcel->removeSheetByIndex($objPHPExcel->getIndex($sheet));
			if($outpath==NULL){
				$outpath = $filepath;
			}
			$writer = PHPExcel_IOFactory::createWriter($objPHPExcel, $excelType);
			$writer->save($outpath);
			return $outpath;
		}
		return NULL;
	}

	/**
	 *	ワークブックのシートを追加
	 *
	 *	@param	$filepath
	 *  @param  $sheetIndex
	 *  @param  $addSheetName
	 *  @param  $addIndex Index where sheet should go (0,1,..., or null for last)
	 *
	 *  		PHPTemplate_ReplaceValue
	 *	@return	array
	 */
	public function addSheet($filepath, $sheetIndex, $addSheetName, $addIndex=NULL, $outpath=NULL, $replaceValues=NULL, $excelType=PHPTemplate_Util::EXCEL_TYPE_2007)
	{
		$reader = PHPExcel_IOFactory::createReader($excelType);
		$objPHPExcel = $reader->load($filepath);

		if (is_int($sheetIndex)) {
			if($objPHPExcel->getSheetCount() > $sheetIndex){
				$sheet = clone $objPHPExcel->getSheet($sheetIndex);
				$sheet->setTitle($addSheetName);
				if($replaceValues!=NULL){
					$sheet = PHPTemplate_Worksheet::replaceValues($sheet, $replaceValues);
				}
				$objPHPExcel->addSheet($sheet, $addIndex);
			}
		}

		if($outpath==NULL){
			$outpath = $filepath;
		}
		$writer = PHPExcel_IOFactory::createWriter($objPHPExcel, $excelType);
		$writer->save($outpath);

		return $outpath;
	}

	/**
	 *	ワークブックのシートを追加
	 *
	 *	@param	$filepath
	 *  @param  $sheetIndex
	 *  @param  $addSheetName
	 *  @param  $addIndex Index where sheet should go (0,1,..., or null for last)
	 *	@return	array
	 */
	public function addSheetByName($filepath, $sheetName, $addSheetName, $addIndex=NULL, $outpath=NULL, $replaceValues=NULL, $excelType=PHPTemplate_Util::EXCEL_TYPE_2007)
	{
		$reader = PHPExcel_IOFactory::createReader($excelType);
		$objPHPExcel = $reader->load($filepath);

		$sheet = $objPHPExcel->getSheetByName($sheetName);
		if($sheet!=NULL){
			$sheet = clone $sheet;
			$sheet->setTitle($addSheetName);
			if($replaceValues!=NULL){
				$sheet = PHPTemplate_Worksheet::replaceValues($sheet, $replaceValues);
			}
			$objPHPExcel->addSheet($sheet, $addIndex);
		}

		if($outpath==NULL){
			$outpath = $filepath;
		}
		$writer = PHPExcel_IOFactory::createWriter($objPHPExcel, $excelType);
		$writer->save($outpath);

		return $outpath;

	}


	/**
	 *	htmlデータテスト用
	 *
	 *	@param	string	  			$filepath
	 *	@param	string			 	$outpath
	 */
	public function htmlTest($filepath, $outpath){
		$reader = PHPExcel_IOFactory::createReader('html');
		$objPHPExcel = $reader->load($filepath);


		// Excel
		$writer = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		$writer->save($outpath);

		return $outpath;
	}

	public function test($filepath, $outpath){
		$reader = PHPExcel_IOFactory::createReader('Excel2007');
		$objPHPExcel = $reader->load($filepath);

		$sheet = $objPHPExcel->getActiveSheet();

		//for($i=0;$i<5000;$i++){
		//	$sheet->insertNewRowBefore(1);
		//}
		$sheet->setSelectedCell();
		$sheet->removeRow(1);

		$writer = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		$writer->save($outpath);

		return $outpath;

	}


}
