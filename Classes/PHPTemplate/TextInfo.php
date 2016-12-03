<?php
/**
 * PHPTemplate_TextInfo
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_TextInfo
 * @copyright
 */
class  PHPTemplate_TextInfo
{
	/**
	 *  タグ付きテキストのタグ名
	 *
	 *  @var  string
	 */
	private $_tagName=NULL;

	/**
	 *  タグ付きテキストのテキスト
	 *
	 *  @var  string
	 */
	private $_text=NULL;

	/**
	 *	Create a new TextInfo
	 *
	 *	@param	string				$tagName
	 *	@param	string				$text
	 *	@throws	PHPExcel_Exception
	 */
	public function __construct($tagName=NULL, $text=NULL)
	{
		$this->_tagName=$tagName;
		$this->_text = $text;
	}

	/**
	 *	Set tag name
	 *
	 *	@param	string				$name
	 */
	public function setTagName($name) {
		$this->_tagName = $name;
	}
	/**
	 *	Set tag text
	 *
	 *	@param	string				$text
	 */
	public function setText($text) {
		$this->_text = $text;
	}

	/**
	 *	Get tag name
	 *
	 *	@return	string
	 */
	public function getTagName() {
		return $this->_tagName;
	}
	/**
	 *	Get tag text
	 *
	 *	@return	string
	 */
	public function getText() {
		return $this->_text;
	}
}

