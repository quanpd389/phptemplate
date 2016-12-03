<?php
/* PHPTemplate_ImportCell
 *
* @category   PHPTemplate
* @package    PHPTemplate_ImportCell
* @copyright
*/
class PHPTemplate_ImportCell
{

	/**
	 *  セルの値
	 *
	 *  @var  string
	 */
	private $_text=NULL;

	/**
	 *  フォント color
	 *
	 *  @var  string
	 */
	private $_fontColor=NULL;

	/**
	 *  セル color
	 *
	 *  @var  string
	 */
	private $_cellColor=NULL;

	/**
	 *  カラム名
	 *
	 *  @var  string
	 */
	private $_columnName=NULL;

	/**
	 *	Create a new ImportCell
	 *
	 *	@param	PHPExcel_Cell $cellr
	 */
	public function __construct($cell, $columnName)
	{
		//$this->_text = PHPTemplate_Cell::getCellText($cell);
		$this->_text = $cell->getCalculatedValue();
		$style = $cell->getStyle();
		$this->_fontColor = $style->getFont()->getColor()->getRGB();
		if($style->getFill()->getFillType() == PHPExcel_Style_Fill::FILL_SOLID){
			$this->_cellColor=$style->getFill()->getStartColor()->getRGB();
		}

		$this->_columnName = $columnName;
//echo($this->_columnName.":".$this->_text."\n");


	}

	/**
	 *	Get column text
	 *
	 *	@return	string
	 */
	public function getText(){
		return $this->_text;
	}

	/**
	 *	Get font color
	 *
	 *	@return	string
	 */
	public function getFontColor(){
		return $this->_fontColor;
	}

	/**
	 *	Get cell color
	 *
	 *	@return	string
	 */
	public function getCellColor(){
		return $this->_cellColor;
	}

	/**
	 *	Get column name
	 *
	 *	@return	string
	 */
	public function getColumnName(){
		return $this->_columnName;
	}
}