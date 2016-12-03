<?php

/* PHPTemplate_SheetConfig
*
* @category   PHPTemplate
* @package    PHPTemplate_SheetConfig
* @copyright
*/
class PHPTemplate_SheetConfig
{
	/**
	 *  シートインデックス
	 *
	 *  @var  int
	 */
	private $_sheetIndex = 1;

	/**
	 *  先頭行を含めるか
	 *
	 *  @var  bool
	 */
	private $_includeFirstRow = false;

	/**
	 *  カラム設定リスト
	 *
	 *  @var  array
	 */
	private $_columnSettingArray = array();

	/**
	 * シートインデックス取得
	 * @return bool
	 */
	public function getSheetIndex(){
		return $this->_sheetIndex;
	}

	/**
	 * 先頭行を含めるか
	 * @return bool
	 */
	public function getIncludeFirstRow(){
		return $this->_includeFirstRow;
	}

	/**
	 * カラム設定リスト取得
	 * @return PHPTemplate_ColumnSetting
	 */
	public function getColumnSetting($idx){
		return $this->_columnSettingArray[$idx];
	}

	/**
	 * カラム数取得
	 *
	 * @return int
	 */
	public function getColumnSettingArrayCount(){
		return count($this->_columnSettingArray);
	}

	/**
	 * 先頭行を含めるか
	 * @param bool $flg
	 */
	public function setIncludeFirstRow($flg){
		$this->_includeFirstRow = $flg;
	}

	/**
	 * カラム情報をセットする
	 * @param SimpleXMLElement $colum
	 */
	public function setColumnSetting($column){
		$columnSetting = new PHPTemplate_ColumnSetting();
		$columnSetting->setColumnSetting($column);

		//$this->_columnSettingArray[(int)($columnSetting->getIndex())] = $columnSetting;
		$this->_columnSettingArray[] = $columnSetting;
	}

	/**
	 * シート情報をセットする
	 * @param SimpleXMLElement $colum
	 */
	public function setSheetSetting($sheet){

		$this->_sheetIndex = (int)($sheet->sheetIndex);

		// 正常に読み込めた場合の処理
		if($sheet->{'include-first-row'} == 'true'){
			$this->_includeFirstRow = true;
		}
		else{
			$this->_includeFirstRow = false;
		}

		foreach ($sheet->columns->column as $val) {
			$this->setColumnSetting($val);
		}
	}



}