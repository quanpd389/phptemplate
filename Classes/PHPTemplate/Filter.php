<?php
/* PHPTemplate_Filter
 *
* @category   PHPTemplate
* @package    PHPTemplate_Filter
* @copyright
*/

class PHPTemplate_FilterColumn
{
	/**
	*  カラム名
	*
	*  @var  string
	*/
	private $_name = NULL;

	/**
	 *  フィルタータイプ
	 * PHPTemplate_Util::FILTER_TYPE_TEXT       = 0;	// 文字データ
	 * PHPTemplate_Util::FILTER_TYPE_FONT_COLOR = 1;	// 文字色
	 * PHPTemplate_Util::FILTER_TYPE_CELL_COLOR = 2;	// セル色
	 *
	 *  @var  int
	 */
	private $_type = PHPTemplate_Util::FILTER_TYPE_TEXT;

	/**
	 *  フィルター値
	 *
	 *  @var  string
	 */
	private $_filter=NULL;

	/**
	 *  セルの値
	 *
	 *  @var  string
	 */
	//private $_text=NULL;

	/**
	 *  フォント color
	 *
	 *  @var  string
	 */
	//private $_fontColor=NULL;

	/**
	 *  セル color
	 *
	 *  @var  string
	 */
	//private $_cellColor=NULL;

	/**
	 *	Create a new Filter
	 *
	 *
	*/
	public function __construct($columnName, $type, $filter)
	{
		$this->_name = $columnName;
		$this->_type = $type;
		$this->_filter = $filter;
	}

	/**
	 * カラム名取得
	 *
	 * @return string
	 */
	public function getName(){
		return $this->_name;
	}

	/**
	 * フィルタータイプ取得
	 *
	 * @return int
	 */
	public function getType(){
		return $this->_type;
	}

	/**
	 * フィルター取得
	 *
	 * @return string
	 */
	public function getFilter(){
		return $this->_filter;
	}

}

class PHPTemplate_Filter
{

	/**
	 *  シート毎のフィルター条件
	 *
	 *  @var  Array
	 */
	private $_sheetFilter = Array();


	/**
	 *	Create a new Filter
	 *
	 *
	 */
	public function __construct()
	{

	}

	/**
	 * シート数取得
	 *
	 * @return int
	 */
	public function getSheetFilterCount(){
		return count($this->_sheetFilter);
	}

	/**
	 * シート毎のフィルター取得
	 * @return Array
	 */
	public function getSheetFilter($idx){
		if (array_key_exists($idx, $this->_sheetFilter)){
			return $this->_sheetFilter[$idx];
		}
		return NULL;
	}

	/**
	 * カラム条件設定
	 * @param string							シートインデックス
	 * @param Array(PHPTemplate_FilterColumn)	フィルター
	 */
	public function setColumnFilter($sheetIndex, $filter)
	{
		$filterArray = Array();
		if(gettype($filter)=='array'){
			$filterArray = $filter;
		}
		else{
			$filterArray[] = $filter;
		}

		if (array_key_exists($sheetIndex, $this->_sheetFilter)){

			$this->_sheetFilter[$sheetIndex][] = $filterArray;

		}
		else{
			$this->_sheetFilter[$sheetIndex] = Array();
			$this->_sheetFilter[$sheetIndex][] = $filterArray;
		}

	}

}