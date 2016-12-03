<?php
/**
 * PHPTemplate_PageFooter
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_PageFooter
 * @copyright
 */
class PHPTemplate_PageFooter extends PHPTemplate_ControlBase
{
	/**
	 *  展開済みのページ番号
	 *
	 *  @var  int
	 */
	private $_page=0;

	/**
	 *	Create a new PageFooter
	 *
	 *	@param	string					$text	parameter
	 *	@param	PHPExcel_Worksheet		$sheet
	 */
	public function __construct($text, $sheet)
	{
		$this->_sheet = $sheet;

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
		if($page->getPagenum() != $this->_page){
			$this->_page = $page->getPagenum();
			parent::expand($data, $rownum, $page, $dataF);
		}

	}

}