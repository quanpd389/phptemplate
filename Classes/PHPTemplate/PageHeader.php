<?php
/**
 * PHPTemplate_PageHeader
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_PageHeader
 * @copyright
 */
class PHPTemplate_PageHeader extends PHPTemplate_ControlBase
{
	/**
	 *	Create a new PageHeader
	 *
	 *	@param	string					$text	parameter
	 *	@param	PHPExcel_Worksheet		$sheet
	 */
	public function __construct($text, $sheet)
	{
		$this->_sheet = $sheet;

	}


}