<?php
/**
 * PHPTemplate_Resume
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_Resume
 * @copyright
 */
class PHPTemplate_Resume extends PHPTemplate_ControlBase
{
	/**
	 *  評価する変数名
	 *
	 *  @var  string
	 */
	private $_value=NULL;

	/**
	 *	Create a new Resume
	 *
	 *	@param	string					$text	parameter
	 *	@param	PHPExcel_Worksheet		$sheet
	 */
	public function __construct($text, $sheet)
	{
		$this->_sheet = $sheet;
		$pos = strpos($text, "#resume");
		if($pos!== false){
			$this->_value = trim(substr($text, strlen("#resume")));
		}
		else{
			throw new PHPTemplate_Exception('#resume parameter is invalid');
		}
	}

	/**
	 * 制御ブロックの展開
	 *
	 * @param array				$data
	 * @param int				$rownum
	 * @param PHPTemplate_Page	$page
	 * @param bool				$dataF
	 */
	public function expand(&$data, &$rownum, &$page, $dataF=TRUE)
	{
		$suspend = $page->getSuspend();
		if(array_key_exists($this->_value, $suspend)){
			$cell = $suspend[$this->_value];
			$cell->bindData($data, $cell->getRow(), $page);
		}
		else{
			throw new PHPTemplate_Exception('#resume parameter cannot find.:'.$this->_value);
		}
	}
}