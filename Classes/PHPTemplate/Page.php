<?php
/**
 * PHPTemplate_Page
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_Page
 * @copyright
 */
class PHPTemplate_Page
{
	/**
	 *  展開中のページ番号
	 *
	 *  @var  int
	 */
	private $_pagenum = 1;
	/**
	 *  ワークシート内のヘッダー
	 *
	 *  @var  PHPTemplate_PageHeader
	 */
	private $_header=NULL;
	/**
	 *  ワークシート内のフッター
	 *
	 *  @var  PHPTemplate_PageFooter
	 */
	private $_footer=NULL;
	/**
	 *  評価遅延指定の変数
	 *
	 *  @var  array
	 */
	private $_suspend=array();
	/**
	 *  ワークシート内でのタグによるフォント指定
	 *
	 *  @var  array
	 */
	private $_font=array();

	/**
	 *	Set page number
	 *
	 *	@param	int				$param
	 */
	public function setPagenum($param) {
		$this->_pagenum = $param;
	}
	/**
	 *	Set header object
	 *
	 *	@param	PHPTemplate_PageHeader		$param
	 */
	public function setHeader($param) {
		$this->_header = $param;
	}
	/**
	 *	Set footer object
	 *
	 *	@param	PHPTemplate_PageFooter		$param
	 */
	public function setFooter($param) {
		$this->_footer = $param;
	}
	/**
	 *	Set suspend value list
	 *
	 *	@param	array		$param
	 */
	public function setSuspend($param) {
		$this->_suspend = $param;
	}
	/**
	 *	Set tag font list
	 *
	 *	@param	array		$param
	 */
	public function setFont($param) {
		$this->_font = $param;
	}
	/**
	 *	Add tag font object
	 *
	 *	@param	PHPTemplate_Font		$param
	 */
	public function addFont($param) {
		$this->_font[] = $param;
	}
	/**
	 *	Add suspend value
	 *
	 *	@param	string		$key
	 *	@param	string		$data
	 */
	public function addSuspend($key, $data) {
		$this->_suspend[$key] = $data;
	}
	/**
	 *	Get page number
	 *
	 *	@return	int
	 */
	public function getPagenum() {
		return $this->_pagenum;
	}
	/**
	 *	Get page header object
	 *
	 *	@return	PHPTemplate_PageHeader
	 */
	public function getHeader() {
		return $this->_header;
	}
	/**
	 *	Get page footer object
	 *
	 *	@return	PHPTemplate_PageFooter
	 */
	public function getFooter() {
		return $this->_footer;
	}
	/**
	 *	Get suspend value list
	 *
	 *	@return	array
	 */
	public function getSuspend() {
		return $this->_suspend;
	}
	/**
	 *	Get tag font list
	 *
	 *	@return	array
	 */
	public function getFont() {
		return $this->_font;
	}

}