<?php
/* PHPTemplate_Filter
 *
* @category   PHPTemplate
* @package    PHPTemplate_Filter
* @copyright
*/

class PHPTemplate_ReplaceValue
{
	/**
	*  置換対象文字列
	*
	*  @var  string
	*/
	private $_searchVal = NULL;

	/**
	 *  置換後文字列
	 *
	 *  @var  string
	 */
	private $_replaceVal = NULL;


	public function __construct($search, $replace)
	{
		$this->_searchVal = $search;
		$this->_replaceVal = $replace;
	}

	/**
	 * 置換対象文字列取得
	 *
	 * @return string
	 */
	public function getSearchVal(){
		return $this->_searchVal;
	}

	/**
	 * 置換後文字列取得
	 *
	 * @return string
	 */
	public function getReplaceVal(){
		return $this->_replaceVal;
	}

}