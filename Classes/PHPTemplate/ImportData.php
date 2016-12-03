<?php
/* PHPTemplate_ImportData
 *
* @category   PHPTemplate
* @package    PHPTemplate_ImportData
* @copyright
*/
class PHPTemplate_ImportData
{
	/**
	 * ワークシート
	 * @var	PHPExcel_Sheet
	 */
	private $_sheet = NULL;

	/**
	 * カラム設定
	 * @var	PHPTemplate_ColumnConfig
	 */
	private $_config = NULL;

	/**
	 * シート内データ
	 * @var	Array
	 */
	private $_rowArrayData = Array();

	/**
	 * シートインデックス
	 * @var	int
	 */
	private $_sheetIndex = 1;


	/**
	 *	Create a new Sheet
	 *
	 *	@param	PHPExcel_Sheet	$sheet
	 */
	public function __construct($sheet, $config, $sheetIndex)
	{
		$this->_sheet = $sheet;
		$this->_config = $config;
		$this->_sheetIndex = $sheetIndex;
	}


	/**
	 * 行数取得
	 *
	 * @return int
	 */
	public function getRowCount(){
		return count($this->_rowArrayData);
	}

	/**
	 * 行データ取得
	 * @return Array
	 */
	public function getRow($idx){
		return $this->_rowArrayData[$idx];
	}

	/**
	 * シートインデックス取得
	 * @return int
	 */
	public function getSheetIndex(){
		return $this->_sheetIndex;
	}

	/**
	 * シートを分析してデータをセット
	 *
	 *
	 */
	public function analyzeSheet(){

		$rowMax = $this->_sheet->getHighestRow();
		$colMax = PHPExcel_Cell::columnIndexFromString($this->_sheet->getHighestColumn());

		$r=1;
		if($this->_config->getIncludeFirstRow()==false){
			$r++;
		}
		for(; $r<=$rowMax; $r++) {
			$rowData = Array();
			for($c = 0; $c < $this->_config->getColumnSettingArrayCount(); $c++){
				$columnSet = $this->_config->getColumnSetting($c);
				$columnIndex = $columnSet->getIndex();
				$columnName = $columnSet->getName();
//echo ($columnName."\n");

				$cell = new PHPTemplate_ImportCell($this->_sheet->getCellByColumnAndRow($columnIndex-1, $r), $columnName);
				$rowData[(string)$columnName] = $cell;
			}
			$this->_rowArrayData[] = $rowData;
		}
	}

	/**
	 * 解析済みデータをフィルタリング
	 *
	 *　@param	PHPTemplate_Filter	$filter
	 */
	public function getFilterData($filter)
	{
		$outRowData = Array();
		//各行のデータを取得
		for($i=0; $i<count($this->_rowArrayData); $i++){
			$row = $this->_rowArrayData[$i];
			if($filter!=NULL){
				foreach($filter as $columnFilterArray){

					$matchFlag=true;
					foreach($columnFilterArray as $columnFilter){
						$columnName = $columnFilter->getName();
						//echo ($columnFilter->getName()."\n");

						$cell = $row[$columnName];

						if($columnFilter->getType() == PHPTemplate_Util::FILTER_TYPE_TEXT){
							if($cell->getText()==$columnFilter->getFilter()){

							}
							else{
								$matchFlag=false;
								break;
							}
						}
						elseif ($columnFilter->getType() == PHPTemplate_Util::FILTER_TYPE_FONT_COLOR){
							if($cell->getFontColor()==$columnFilter->getFilter()){

							}
							else{
								$matchFlag=false;
								break;
							}
						}
						else{
							if($cell->getCellColor()==$columnFilter->getFilter()){

							}
							else{
								$matchFlag=false;
								break;
							}
						}
					}
					if($matchFlag){
						//echo("フィルターあり：".($i+1)."行目取得\n");
						$outData = $this->getRowOutArray($row);
						$outRowData[] = $outData;
					}
				}
			}
			else{
				//echo("フィルターなし：".($i+1)."行目取得\n");
				$outData = PHPTemplate_ImportData::getRowOutArray($row);
				$outRowData[] = $outData;

			}
		}
		return $outRowData;
	}

	/**
	 * 指定行のデータをArrayにして返す
	 *
	 *
	 */
	public function getRowOutArray($row){

		$outArray=Array();
		foreach($row as $cell){
			$key = $cell->getColumnName();
			$outArray[(string)$key] = $cell->getText();
		}
		return $outArray;
	}

}