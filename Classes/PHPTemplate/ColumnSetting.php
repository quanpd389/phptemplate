<?php
/* PHPTemplate_ColumnSetting
 *
* @category   PHPTemplate
* @package    PHPTemplate_ColumnSetting
* @copyright
*/
class PHPTemplate_ColumnSetting
{
	/**
	 *  行番号(1から)
	 *
	 *  @var  int
	 */
	private $_index = 1;

	/**
	 *  カラム名
	 *
	 *  @var  string
	 */
	private $_name;


	/**
	 * カラム番号取得
	 * @return int
	 */
	public function getIndex(){
		return $this->_index;
	}

	/**
	 * カラム名取得
	 * @return string
	 */
	public function getName(){
		return $this->_name;
	}

	/**
	 * カラム番号設定
	 * @param int $idx
	 */
	public function setIndex($idx){
		$this->_index = $idx;
	}

	/**
	 * カラム名設定
	 * @param string $name
	 */
	public function setName($name){
		$this->_name = $name;
	}

	/**
	 * 行番号・カラム名設定
	 * @param SimpleXMLElement  $column
	 */
	public function setColumnSetting($column){
		$this->_index = $column->index;
		$this->_name = $column->name;
	}

}