<?php
/**
 * PHPTemplate_RowMerge
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_RowMerge
 * @copyright
 */
class PHPTemplate_RowMerge
{

	/**
	 *  結合する行数
	 *
	 *  @var  int
	 */
	private $_rowCnt=0;
	/**
	 *  結合するカラム情報
	 *
	 *  @var  string[]
	 */
	private $_mergeInfo=NULL;
	/**
	 *  結合するセルのスタイル
	 *
	 *  @var  PHPExcel_Style
	 */
	private $_style=NULL;

	/**
	 *	Create a new RowMerge
	 *
	 *	@param	int					$cnt	結合する行数
	 *	@param	string[]			$info	結合するカラム情報
	 *	@param	PHPExcel_Style		$style	結合するセルのスタイル
	 */
	public function __construct($cnt=NULL, $info=NULL, $style=NULL)
	{
		$this->_rowCnt=$cnt;
		$this->_mergeInfo = $info;
		$this->_style = $style;
	}
	/**
	 *	Set row count
	 *
	 *	@param		int		$param
	 */
	public function setRowCnt($param) {
		$this->_rowCnt=$param;
	}
	/**
	 *	Set merge info
	 *
	 *	@param		string[]	$param
	 */
	public function setMergeInfo($param) {
		$this->_mergeInfo=$param;
	}
	/**
	 *	Set cell style
	 *
	 *	@param		PHPExcel_Style		$param
	 */
	public function setStyle($param) {
		$this->_style=$param;
	}
	/**
	 *	Get row count
	 *
	 *	@return int
	 */
	public function getRowCnt() {
		return $this->_rowCnt;
	}
	/**
	 *	Get merge info
	 *
	 *	@return string[]
	 */
	public function getMergeInfo() {
		return $this->_mergeInfo;
	}
	/**
	 *	Get cell Style
	 *
	 *	@return PHPExcel_Style
	 */
	public function getStyle() {
		return $this->_style;
	}

}