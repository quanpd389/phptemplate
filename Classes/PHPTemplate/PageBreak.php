<?php
/**
 * PHPTemplate_PageBreak
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_PageBreak
 * @copyright
 */
class PHPTemplate_PageBreak extends PHPTemplate_ControlBase
{
	/**
	 *  ページヘッダー
	 *
	 *  @var  PHPTemplate_PageHeader
	 */
	private $_pageHeader=NULL;
	/**
	 *  ページフッター
	 *
	 *  @var  PHPTemplate_PageFooter
	 */
	private $_pageFooter=NULL;

	/**
	 *	Create a new PageBreak
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
		if($this->_pageHeader==NULL){
			$this->_pageHeader = $page->getHeader();
		}
		if($this->_pageFooter==NULL){
			$this->_pageFooter = $page->getFooter();
		}

		if($this->_pageFooter != NULL){
			$this->_pageFooter->expand($data, $rownum, $page);
		}

		$this->_sheet->setBreak("A".(string)($rownum-1), PHPExcel_Worksheet::BREAK_ROW);
		$page->setPagenum($page->getPagenum() + 1);

		if($this->_pageHeader != NULL){
			$this->_pageHeader->expand($data, $rownum, $page);
		}
	}
}