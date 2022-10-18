<?php

namespace NIJC;

use PhpOffice\PhpWord;

/**
 * abstracts writer class.
 */
class AbstractsWriter
{
    // non-break space (UTF-8)
    public const NBSP = "\xc2\xa0";
    // copyright sign (UTF-8)
    public const COPYRIGHT_SIGN = "\xc2\xa9";
    // figure dpi
    public const FIGURE_DPI = 96;
    // font size of texvc generated image.
    public const TEXVC_FONT_SIZE = 14;
    // font
    public const FONT_SANS = 'Calibri';
    public const FONT_SANS_SERIF = 'Cambria';
    // heading labes
    public const HL_ABSTRACTS = 'Abstracts';
    public const HL_DOI = 'doi';
    public const HL_TOC = 'Table of Contents';
    public const HL_FIGURE_CAPTION = 'Figure';
    public const HL_ACKNOWLEDGEMENTS = 'Acknowledgements';
    public const HL_REFERENCES = 'References';
    public const HL_TOPIC = 'Topic';
    public const HL_CITATION = 'Citation';
    public const HL_COPYRIGHT = 'Copyright';
    public const HL_EDITED_BY = 'Edited by';
    public const HL_PUBLISHED_BY = 'Published by';
    // title style depth.
    public const TS_SESSION_TITLE = 1;
    public const TS_ARTICLE_TITLE = 2;
    // numbering style name.
    public const NS_ARTICLE_REFERENCES = 'NS_Article_References';
    // font style name.
    public const FS_TITLEPAGE_BOOK_TITLE = 'FS_TitlePage_BookTitle';
    public const FS_TITLEPAGE_BOOK_DOI = 'FS_TitlePage_BookDOI';
    public const FS_TITLEPAGE_CONFERENCE_TITLE = 'FS_TitlePage_ConferenceTitle';
    public const FS_HEADER = 'FS_Header';
    public const FS_FOOTER = 'FS_Footer';
    public const FS_TOC_TITLE = 'FS_TOC_Title';
    public const FS_ARTICLE_TITLE = 'FS_Article_Title';
    public const FS_ARTICLE_AUTHOR = 'FS_Article_Author';
    public const FS_ARTICLE_AUTHOR_SUPER = 'FS_Article_AuthorSuper';
    public const FS_ARTICLE_AFFILIATION = 'FS_Article_Affiliation';
    public const FS_ARTICLE_AFFILIATION_SUPER = 'FS_Article_AffiliationSuper';
    public const FS_ARTICLE_EMAIL = 'FS_Article_Email';
    public const FS_ARTICLE_DOI = 'FS_Article_DOI';
    public const FS_ARTICLE_ABSTRACT = 'FS_Article_Abstract';
    public const FS_ARTICLE_FIGURE_CAPTION_HEADING = 'FS_Article_FigureCaptionHeading';
    public const FS_ARTICLE_FIGURE_CAPTION = 'FS_Article_FigureCaption';
    public const FS_ARTICLE_ACKNOWLEDGEMENTS_HEADING = 'FS_Article_AcknowlegedmentsHeading';
    public const FS_ARTICLE_ACKNOWLEDGEMENTS = 'FS_Article_Acknowlegedments';
    public const FS_ARTICLE_REFERENCES_HEADING = 'FS_Article_ReferencesHeading';
    public const FS_ARTICLE_REFERENCES = 'FS_Article_References';
    public const FS_ARTICLE_TOPIC_HEADING = 'FS_Article_TopicHeading';
    public const FS_ARTICLE_TOPIC = 'FS_Article_Topic';
    public const FS_ARTICLE_CITATION_HEADING = 'FS_Article_CitationHeading';
    public const FS_ARTICLE_CITATION = 'FS_Article_Citation';
    public const FS_ARTICLE_COPYRIGHT_HEADING = 'FS_Article_CopyrightHeading';
    public const FS_ARTICLE_COPYRIGHT = 'FS_Article_Copyright';
    public const FS_COLOPHON_BOOK_TITLE = 'FS_Colophon_BookTitle';
    public const FS_COLOPHON_CONFERENCE_TITLE = 'FS_Colophon_ConferenceTitle';
    public const FS_COLOPHON_PUBLICATION_DATE = 'FS_Colophon_PublicationDate';
    public const FS_COLOPHON_EDITOR_HEADING = 'FS_Colophon_EditorHeading';
    public const FS_COLOPHON_EDITOR = 'FS_Colophon_Editor';
    public const FS_COLOPHON_PUBLISHER_HEADING = 'FS_Colophon_PublisherHeading';
    public const FS_COLOPHON_PUBLISHER = 'FS_Colophon_Publisher';
    public const FS_COLOPHON_COPYRIGHT = 'FS_Colophon_Copyright';
    // paragraph style name.
    public const PS_TITLEPAGE_BOOK_TITLE = 'PS_TitlePage_BookTitle';
    public const PS_TITLEPAGE_BOOK_DOI = 'PS_TitlePage_BookDOI';
    public const PS_TITLEPAGE_CONFERENCE_TITLE = 'PS_TitlePage_ConferenceTitle';
    public const PS_HEADER = 'PS_Header';
    public const PS_FOOTER = 'PS_Footer';
    public const PS_TOC_TITLE = 'PS_TOC_Title';
    public const PS_ARTICLE_TITLE = 'PS_Article_Title';
    public const PS_ARTICLE_AUTHOR = 'PS_Article_Author';
    public const PS_ARTICLE_AFFILIATION = 'PS_Article_Affiliation';
    public const PS_ARTICLE_EMAIL = 'PS_Article_Email';
    public const PS_ARTICLE_DOI = 'PS_Article_DOI';
    public const PS_ARTICLE_ABSTRACT = 'PS_Article_Abstract';
    public const PS_ARTICLE_ABSTRACT_FOLLOW = 'PS_Article_AbstractFollow';
    public const PS_ARTICLE_FIGURE = 'PS_Article_Figure';
    public const PS_ARTICLE_FIGURE_CAPTION = 'PS_Article_FigureCaption';
    public const PS_ARTICLE_ACKNOWLEDGEMENTS_HEADING = 'PS_Article_AcknowlegedmentsHeading';
    public const PS_ARTICLE_ACKNOWLEDGEMENTS = 'PS_Article_Acknowlegedments';
    public const PS_ARTICLE_REFERENCES_HEADING = 'PS_Article_ReferencesHeading';
    public const PS_ARTICLE_REFERENCES = 'PS_Article_References';
    public const PS_ARTICLE_TOPIC = 'PS_Article_Topic';
    public const PS_ARTICLE_CITATION = 'PS_Article_Citation';
    public const PS_ARTICLE_COPYRIGHT = 'PS_Article_Copyright';
    public const PS_COLOPHON_BOOK_TITLE = 'PS_Colophon_BookTitle';
    public const PS_COLOPHON_CONFERENCE_TITLE = 'PS_Colophon_ConferenceTitle';
    public const PS_COLOPHON_PUBLICATION_DATE = 'PS_Colophon_PublicationDate';
    public const PS_COLOPHON_EDITOR_HEADING = 'PS_Colophon_EditorHeading';
    public const PS_COLOPHON_EDITOR = 'PS_Colophon_Editor';
    public const PS_COLOPHON_PUBLISHER_HEADING = 'PS_Colophon_PublisherHeading';
    public const PS_COLOPHON_PUBLISHER = 'PS_Colophon_Publisher';
    public const PS_COLOPHON_COPYRIGHT = 'PS_Colophon_Copyright';

    /**
     * abstracts.
     *
     * @var array
     */
    protected $mAbstracts;

    /**
     * conference.
     *
     * @var array
     */
    protected $mConference;

    /**
     * book.
     *
     * @var array
     */
    protected $mBook;

    /**
     * images directory.
     *
     * @var string
     */
    protected $mImagesDir;

    /**
     * images directory.
     *
     * @var \NIJC\TexvcCommand
     */
    protected $mTexvc;

    /**
     * default font name.
     *
     * @var string
     */
    protected $mDefaultFontName;

    /**
     * default font size.
     *
     * @var int
     */
    protected $mDefaultFontSize;

    /**
     * maximum figure width (mm).
     *
     * @var int
     */
    protected $mMaxFigureWidth;

    /**
     * maximum figure height (mm).
     *
     * @var int
     */
    protected $mMaxFigureHeight;

    /**
     * section style.
     *
     * @var array
     */
    protected $mSectionStyle;

    /**
     * title styles.
     *
     * @var array
     */
    protected $mTitleStyles;

    /**
     * numbering styles.
     *
     * @var array
     */
    protected $mNumberingStyles;

    /**
     * font styles.
     *
     * @var array
     */
    protected $mFontStyles;

    /**
     * paragraph styles.
     *
     * @var array
     */
    protected $mParagraphStyles;

    /**
     * constructor.
     *
     * @param string $ajson
     * @param string $imagesDir
     */
    public function __construct($ajson, $cjson, $imagesDir)
    {
        $this->mAbstracts = json_decode($ajson, true);
        if (null === $this->mAbstracts) {
            exit('Fatal Error: failed to decode abstracts json data'.PHP_EOL);
        }
        $this->mConference = json_decode($cjson, true);
        if (null === $this->mConference) {
            exit('Fatal Error: failed to decode conference json data'.PHP_EOL);
        }
        $bjsonPath = dirname(__DIR__).'/config/conference/'.$this->mConference['short'].'.json';
        if (!file_exists($bjsonPath)) {
            $bjsonPath = dirname(__DIR__).'/config/conference/default.json';
            if (!file_exists($bjsonPath)) {
                exit('Fatal Error: book json file not found: '.$bjsonPath.PHP_EOL);
            }
        }
        $bjson = file_get_contents($bjsonPath);
        $this->mBook = json_decode($bjson, true);
        if (null === $this->mBook) {
            exit('Fatal Error: failed to decode book json data'.PHP_EOL);
        }
        $this->mImagesDir = $imagesDir;
        $this->mTexvc = new TexvcCommand($imagesDir);
        $this->_initStyles();
    }

    /**
     * save word file.
     *
     * @param string $filePath
     *
     * @return bool
     */
    public function save($filePath)
    {
        $phpword = $this->_createDocument();
        $section = $this->_addSection($phpword);
        $this->_printInfo('Title Page');
        $this->_addTitlePage($section);
        $section->addPageBreak();
        $section = $this->_addSection($phpword);
        $this->_addHeader($section);
        $this->_addFooter($section);
        // Recent version of PhpWord doesn't support image contained title for the TOC output.
        // $this->_printInfo('Table of Contents');
        // $this->_addTableOfContentsPage($section);
        $numArticles = 0;
        $sessionTypeLast = '';
        foreach ($this->mAbstracts as $abstract) {
            if ('Accepted' !== $abstract['state']) {
                continue;
            }
            ++$numArticles;
            $sessionType = $abstract['abstrTypes'][0]['short'] ?? '';
            if ($sessionType !== $sessionTypeLast) {
                $sessionTitle = $this->_getSessionTitle($sessionType);
                $sessionDescription = $this->_getSessionDescription($sessionType);
                $section = $this->_addSection($phpword);
                $this->_printInfo('Section Cover Page - '.$sessionTitle);
                $this->_addSessionCoverPage($section, $sessionTitle, $sessionDescription);
                $sessionTypeLast = $sessionType;
            }
            $section->addPageBreak();
            $section = $this->_addSection($phpword);
            $phpword->addNumberingStyle(self::NS_ARTICLE_REFERENCES.$numArticles, $this->mNumberingStyles[self::NS_ARTICLE_REFERENCES]);
            $this->_printInfo('Article - '.$abstract['title']);
            $this->_addArticlePage($section, $abstract, $numArticles);
        }
        $section->addPageBreak();
        $section = $this->_addSection($phpword);
        $this->_printInfo('Colophan');
        $this->_addColophonPage($section);
        $writer = PhpWord\IOFactory::createWriter($phpword, 'Word2007');
        $writer->save($filePath);

        return true;
    }

    /**
     * get abstract book title.
     *
     * @return string
     */
    protected function _getBookTitle()
    {
        return $this->mBook['title'];
    }

    /**
     * get abstract book doi.
     *
     * @return string
     */
    protected function _getBookDOI()
    {
        return $this->mBook['doi'];
    }

    /**
     * get conference title.
     *
     * @return string
     */
    protected function _getConferenceTitle()
    {
        return $this->mBook['conference']['title'];
    }

    /**
     * get session title from abstract type(sort).
     *
     * @param string $type
     *
     * @return string
     */
    protected function _getSessionTitle($type)
    {
        foreach ($this->mBook['session'] as $session) {
            if ($session['short'] == $type) {
                return $session['title'];
            }
        }
        foreach ($this->mConference['groups'] as $session) {
            if ($session['short'] == $type) {
                return $session['name'];
            }
        }

        return null;
    }

    /**
     * get session description from abstract type(sort).
     *
     * @param string $type
     *
     * @return string
     */
    protected function _getSessionDescription($type)
    {
        foreach ($this->mBook['session'] as $session) {
            if ($session['short'] == $type) {
                return $session['description'];
            }
        }

        return null;
    }

    /**
     * get publication year.
     *
     * @return string
     */
    protected function _getPublicationYear()
    {
        $ret = isset($this->mBook['publicationDate']['year']) ? $this->mBook['publicationDate']['year'] : null;

        return null === $ret ? '' : trim($ret);
    }

    /**
     * get publication month.
     *
     * @return string
     */
    protected function _getPublicationMonth()
    {
        $ret = isset($this->mBook['publicationDate']['month']) ? $this->mBook['publicationDate']['month'] : null;

        return null === $ret ? '' : trim($ret);
    }

    /**
     * get editor.
     *
     * @return array
     */
    protected function _getEditor()
    {
        $ret = isset($this->mBook['editor']) ? $this->mBook['editor'] : null;

        return null === $ret ? [] : array_map('trim', $ret);
    }

    /**
     * get publisher name.
     *
     * @return string
     */
    protected function _getPublisherName()
    {
        $ret = isset($this->mBook['publisher']['name']) ? $this->mBook['publisher']['name'] : null;

        return null === $ret ? '' : trim($ret);
    }

    /**
     * get publisher address.
     *
     * @return string
     */
    protected function _getPublisherAddress()
    {
        $ret = isset($this->mBook['publisher']['address']) ? $this->mBook['publisher']['address'] : null;

        return null === $ret ? '' : trim($ret);
    }

    /**
     * initialize styles.
     */
    protected function _initStyles()
    {
        $this->mDefaultFontName = self::FONT_SANS;
        $this->mDefaultFontSize = 10;
        $this->mMaxFigureWidth = 140;
        $this->mMaxFigureHeight = 75;
        $this->mSectionStyle = [
            'paperSize' => 'A4',
            'orientation' => 'portrait',
            'marginTop' => PhpWord\Shared\Converter::cmToTwip(2),
            'marginRight' => PhpWord\Shared\Converter::cmToTwip(2),
            'marginLeft' => PhpWord\Shared\Converter::cmToTwip(2),
            'marginBottom' => PhpWord\Shared\Converter::cmToTwip(1.5),
        ];
        $this->mTitleStyles = [
            self::TS_SESSION_TITLE => [
                [
                    'name' => self::FONT_SANS_SERIF,
                    'size' => 16,
                    'bold' => true,
                    'italic' => true,
                ], [
                    'alignment' => PhpWord\SimpleType\Jc::CENTER,
                    'lineHeight' => 1.5,
                    'space' => [
                        'before' => PhpWord\Shared\Converter::cmToTwip(9),
                        'after' => PhpWord\Shared\Converter::pointToTwip(16),
                    ],
                ],
            ],
            self::TS_ARTICLE_TITLE => [
                [
                    'size' => 14,
                    'bold' => true,
                ], [
                    'alignment' => PhpWord\SimpleType\Jc::CENTER,
                    'space' => [
                        'before' => PhpWord\Shared\Converter::cmToTwip(0.6),
                        'after' => PhpWord\Shared\Converter::pointToTwip(1),
                    ],
                ],
            ],
        ];
        $this->mNumberingStyles = [
            self::NS_ARTICLE_REFERENCES => [
                'type' => 'singleLevel',
                'levels' => [
                    [
                        'format' => 'decimal',
                        'text' => '%1.',
                        'left' => 360,
                        'hanging' => 360,
                        'tabPos' => 360,
                    ],
                ],
            ],
        ];
        $this->mFontStyles = [
            self::FS_TITLEPAGE_BOOK_TITLE => [
                'size' => 22,
                'bold' => true,
            ],
            self::FS_TITLEPAGE_BOOK_DOI => [
                'size' => 18,
                'bold' => true,
            ],
            self::FS_TITLEPAGE_CONFERENCE_TITLE => [
                'name' => self::FONT_SANS_SERIF,
                'size' => 18,
                'bold' => true,
                'italic' => true,
            ],
            self::FS_HEADER => [
                'size' => 11,
                'italic' => true,
            ],
            self::FS_FOOTER => [
                'size' => 11,
            ],
            self::FS_TOC_TITLE => [
                'size' => 18,
                'bold' => true,
            ],
            self::FS_ARTICLE_TITLE => $this->mTitleStyles[self::TS_ARTICLE_TITLE][0],
            self::FS_ARTICLE_AUTHOR => [
                'size' => 10.5,
                'bold' => true,
            ],
            self::FS_ARTICLE_AUTHOR_SUPER => [
                'size' => 10.5,
                'bold' => true,
                'superScript' => true,
            ],
            self::FS_ARTICLE_AFFILIATION => [
                'size' => 9,
                'italic' => true,
            ],
            self::FS_ARTICLE_AFFILIATION_SUPER => [
                'size' => 9,
                'italic' => true,
                'superScript' => true,
            ],
            self::FS_ARTICLE_EMAIL => [
            ],
            self::FS_ARTICLE_DOI => [
                'size' => 10.5,
            ],
            self::FS_ARTICLE_ABSTRACT => [
                'name' => self::FONT_SANS_SERIF,
            ],
            self::FS_ARTICLE_FIGURE_CAPTION_HEADING => [
                'bold' => true,
            ],
            self::FS_ARTICLE_FIGURE_CAPTION => [
            ],
            self::FS_ARTICLE_ACKNOWLEDGEMENTS_HEADING => [
                'size' => 10.5,
                'bold' => true,
            ],
            self::FS_ARTICLE_ACKNOWLEDGEMENTS => [
            ],
            self::FS_ARTICLE_REFERENCES_HEADING => [
                'size' => 10.5,
                'bold' => true,
            ],
            self::FS_ARTICLE_REFERENCES => [
            ],
            self::FS_ARTICLE_TOPIC_HEADING => [
                'bold' => true,
            ],
            self::FS_ARTICLE_TOPIC => [
            ],
            self::FS_ARTICLE_CITATION_HEADING => [
                'bold' => true,
            ],
            self::FS_ARTICLE_CITATION => [
            ],
            self::FS_ARTICLE_COPYRIGHT_HEADING => [
                'bold' => true,
            ],
            self::FS_ARTICLE_COPYRIGHT => [
            ],
            self::FS_COLOPHON_BOOK_TITLE => [
                'bold' => true,
                'size' => 14,
            ],
            self::FS_COLOPHON_CONFERENCE_TITLE => [
                'name' => self::FONT_SANS_SERIF,
                'bold' => true,
                'italic' => true,
                'size' => 14,
            ],
            self::FS_COLOPHON_PUBLICATION_DATE => [
                'size' => 10.5,
            ],
            self::FS_COLOPHON_EDITOR_HEADING => [
                'bold' => true,
                'size' => 10.5,
            ],
            self::FS_COLOPHON_EDITOR => [
                'size' => 10.5,
            ],
            self::FS_COLOPHON_PUBLISHER_HEADING => [
                'bold' => true,
                'size' => 10.5,
            ],
            self::FS_COLOPHON_PUBLISHER => [
                'size' => 10.5,
            ],
            self::FS_COLOPHON_COPYRIGHT => [
                'size' => 10.5,
            ],
        ];
        $this->mParagraphStyles = [
            self::PS_TITLEPAGE_BOOK_TITLE => [
                'alignment' => PhpWord\SimpleType\Jc::CENTER,
                'space' => [
                    'before' => PhpWord\Shared\Converter::cmToTwip(6),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_TITLEPAGE_BOOK_DOI => [
                'alignment' => PhpWord\SimpleType\Jc::CENTER,
                'space' => [
                    'before' => PhpWord\Shared\Converter::cmToTwip(1.5),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_TITLEPAGE_CONFERENCE_TITLE => [
                'alignment' => PhpWord\SimpleType\Jc::CENTER,
                'space' => [
                    'before' => PhpWord\Shared\Converter::cmToTwip(2),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_HEADER => [
                'alignment' => PhpWord\SimpleType\Jc::RIGHT,
                'borderBottomSize' => 1,
            ],
            self::PS_FOOTER => [
                'alignment' => PhpWord\SimpleType\Jc::CENTER,
                'borderTopSize' => 1,
            ],
            self::PS_TOC_TITLE => [
                'alignment' => PhpWord\SimpleType\Jc::CENTER,
                'space' => [
                   'before' => PhpWord\Shared\Converter::pointToTwip(30),
                   'after' => PhpWord\Shared\Converter::pointToTwip(10),
                ],
            ],
            self::PS_ARTICLE_TITLE => $this->mTitleStyles[self::TS_ARTICLE_TITLE][1],
            self::PS_ARTICLE_AUTHOR => [
                'alignment' => PhpWord\SimpleType\Jc::CENTER,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(5),
                    'after' => PhpWord\Shared\Converter::pointToTwip(5),
                ],
            ],
            self::PS_ARTICLE_AFFILIATION => [
                'alignment' => PhpWord\SimpleType\Jc::CENTER,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(3),
                    'after' => PhpWord\Shared\Converter::pointToTwip(3),
                ],
            ],
            self::PS_ARTICLE_EMAIL => [
                'alignment' => PhpWord\SimpleType\Jc::CENTER,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(3),
                    'after' => PhpWord\Shared\Converter::pointToTwip(3),
                ],
            ],
            self::PS_ARTICLE_DOI => [
                'alignment' => PhpWord\SimpleType\Jc::CENTER,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(5),
                    'after' => PhpWord\Shared\Converter::pointToTwip(5),
                ],
            ],
            self::PS_ARTICLE_ABSTRACT => [
                'alignment' => PhpWord\SimpleType\Jc::BOTH,
                'lineHeight' => 1.05,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(5),
                    'after' => PhpWord\Shared\Converter::pointToTwip(3),
                ],
            ],
            self::PS_ARTICLE_ABSTRACT_FOLLOW => [
                'alignment' => PhpWord\SimpleType\Jc::BOTH,
                'indentation' => [
                    'firstLine' => PhpWord\Shared\Converter::cmToTwip(0.4),
                ],
                'lineHeight' => 1.05,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(1),
                    'after' => PhpWord\Shared\Converter::pointToTwip(3),
                ],
            ],
            self::PS_ARTICLE_FIGURE => [
                'alignment' => PhpWord\SimpleType\Jc::CENTER,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(10),
                    'after' => PhpWord\Shared\Converter::pointToTwip(10),
                ],
            ],
            self::PS_ARTICLE_FIGURE_CAPTION => [
                'alignment' => PhpWord\SimpleType\Jc::BOTH,
                'indentation' => [
                    'left' => PhpWord\Shared\Converter::cmToTwip(2),
                    'right' => PhpWord\Shared\Converter::cmToTwip(2),
                ],
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(3),
                    'after' => PhpWord\Shared\Converter::pointToTwip(10),
                ],
            ],
            self::PS_ARTICLE_ACKNOWLEDGEMENTS_HEADING => [
                'alignment' => PhpWord\SimpleType\Jc::LEFT,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(10),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_ARTICLE_ACKNOWLEDGEMENTS => [
                'alignment' => PhpWord\SimpleType\Jc::BOTH,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(3),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_ARTICLE_REFERENCES_HEADING => [
                'alignment' => PhpWord\SimpleType\Jc::LEFT,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(10),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_ARTICLE_REFERENCES => [
                'alignment' => PhpWord\SimpleType\Jc::BOTH,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(1),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_ARTICLE_TOPIC => [
                'alignment' => PhpWord\SimpleType\Jc::BOTH,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(10),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_ARTICLE_CITATION => [
                'alignment' => PhpWord\SimpleType\Jc::BOTH,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(10),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_ARTICLE_COPYRIGHT => [
                'alignment' => PhpWord\SimpleType\Jc::BOTH,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(10),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_COLOPHON_BOOK_TITLE => [
                'alignment' => PhpWord\SimpleType\Jc::LEFT,
                'space' => [
                    'before' => PhpWord\Shared\Converter::cmToTwip(8),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_COLOPHON_CONFERENCE_TITLE => [
                'alignment' => PhpWord\SimpleType\Jc::LEFT,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(10),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_COLOPHON_PUBLICATION_DATE => [
                'alignment' => PhpWord\SimpleType\Jc::LEFT,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(10),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_COLOPHON_EDITOR_HEADING => [
                'alignment' => PhpWord\SimpleType\Jc::LEFT,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(15),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_COLOPHON_EDITOR => [
                'alignment' => PhpWord\SimpleType\Jc::LEFT,
                'indentation' => [
                    'left' => PhpWord\Shared\Converter::pointToTwip(15),
                ],
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(3),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_COLOPHON_PUBLISHER_HEADING => [
                'alignment' => PhpWord\SimpleType\Jc::LEFT,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(15),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_COLOPHON_PUBLISHER => [
                'alignment' => PhpWord\SimpleType\Jc::LEFT,
                'indentation' => [
                    'left' => PhpWord\Shared\Converter::pointToTwip(15),
                ],
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(3),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
            self::PS_COLOPHON_COPYRIGHT => [
                'alignment' => PhpWord\SimpleType\Jc::LEFT,
                'space' => [
                    'before' => PhpWord\Shared\Converter::pointToTwip(15),
                    'after' => PhpWord\Shared\Converter::pointToTwip(1),
                ],
            ],
        ];
    }

    /**
     * create document.
     *
     * @return \PhpOffice\PhpWord\PhpWord
     */
    protected function _createDocument()
    {
        $phpword = new PhpWord\PhpWord();
        $phpword->setDefaultFontName($this->mDefaultFontName);
        $phpword->setDefaultFontSize($this->mDefaultFontSize);
        foreach ($this->mTitleStyles as $level => $styles) {
            list($fontStyle, $paragraphStyle) = $styles;
            $phpword->addTitleStyle($level, $fontStyle, $paragraphStyle);
        }
        foreach ($this->mFontStyles as $name => $style) {
            $phpword->addFontStyle($name, $style);
        }
        foreach ($this->mParagraphStyles as $name => $style) {
            $phpword->addParagraphStyle($name, $style);
        }

        return $phpword;
    }

    /**
     * add section.
     *
     * @param \PhpOffice\PhpWord\PhpWord $phpword
     */
    protected function _addSection($phpword)
    {
        return $phpword->addSection($this->mSectionStyle);
    }

    /**
     * add header.
     *
     * @param \PhpOffice\PhpWord\Element\Section $section
     */
    protected function _addHeader($section)
    {
        $header = $section->createHeader();
        $header->addText($this->_getBookTitle(), self::FS_HEADER, self::PS_HEADER);
    }

    /**
     * add footer.
     *
     * @param \PhpOffice\PhpWord\Element\Section $section
     */
    protected function _addFooter($section)
    {
        $footer = $section->createFooter();
        $footer->addPreserveText('{PAGE}', self::FS_FOOTER, self::PS_FOOTER);
    }

    /**
     * add title page.
     *
     * @param \PhpOffice\PhpWord\Element\Section $section
     */
    protected function _addTitlePage($section)
    {
        // book title
        $textArr = explode("\n", $this->_getBookTitle());
        $run = $section->addTextRun(self::PS_TITLEPAGE_BOOK_TITLE);
        foreach ($textArr as $num => $text) {
            if (0 < $num) {
                $run->addTextBreak();
            }
            $run->addText($this->_sanitizeText($text), self::FS_TITLEPAGE_BOOK_TITLE);
        }
        // conference title
        $textArr = explode("\n", $this->_getConferenceTitle());
        $run = $section->addTextRun(self::PS_TITLEPAGE_CONFERENCE_TITLE);
        foreach ($textArr as $num => $text) {
            if (0 < $num) {
                $run->addTextBreak();
            }
            $run->addText($this->_sanitizeText($text), self::FS_TITLEPAGE_CONFERENCE_TITLE);
        }
        $run->addTextBreak();
        $run->addText(self::HL_ABSTRACTS, self::FS_TITLEPAGE_CONFERENCE_TITLE);
        // DOI
        $doi = $this->_getBookDOI();
        if ($doi) {
            $text = self::HL_DOI.':'.$this->_sanitizeText($doi);
            $section->addText($text, self::FS_TITLEPAGE_BOOK_DOI, self::PS_TITLEPAGE_BOOK_DOI);
        }
    }

    /**
     * add table of contents page.
     *
     * @param \PhpOffice\PhpWord\Element\Section $section
     */
    protected function _addTableOfContentsPage($section)
    {
        $section->addText(self::HL_TOC, self::FS_TOC_TITLE, self::PS_TOC_TITLE);
        $section->addTOC();
    }

    /**
     * add session cover page.
     *
     * @param \PhpOffice\PhpWord\Element\Section $section
     * @param string                             $name
     * @param string                             $description
     */
    protected function _addSessionCoverPage($section, $name, $description)
    {
        $text = $name;
        if (!empty($description)) {
            $text .= "\n".$description;
        }
        $text = preg_replace('/\R/u', '</w:t><w:br/><w:t>', $this->_sanitizeText($text));
        $section->addTitle($text, self::TS_SESSION_TITLE);
    }

    /**
     * add article page.
     *
     * @param \PhpOffice\PhpWord\Element\Section $section
     * @param array                              $abstract
     * @param int                                $numArticles
     */
    protected function _addArticlePage($section, $abstract, $numArticles)
    {
        // Title
        $title = $this->_sanitizeText($abstract['title']);
        if (preg_match('/\$(.+)\$/', $title)) {
            $run = new PhpWord\Element\TextRun(self::PS_ARTICLE_TITLE);
            $this->_addRunTextWithMath($run, $title, self::FS_ARTICLE_TITLE);
            $title = $run;
        }
        $section->addTitle($title, self::TS_ARTICLE_TITLE);
        // Authors
        $run = $section->addTextRun(self::PS_ARTICLE_AUTHOR);
        foreach ($abstract['authors'] as $authorIdx => $author) {
            if ($authorIdx !== $author['position']) {
                exit('Fatal Error: "authors:position" is not sorted (uuid:'.$author['uuid'].')'.PHP_EOL);
            }
            if (0 < $authorIdx) {
                $run->addText(', ', self::FS_ARTICLE_AUTHOR);
            }
            // name
            $firstName = isset($author['firstName']) ? $this->_sanitizeText($author['firstName']) : '';
            $middleName = isset($author['middleName']) ? $this->_sanitizeText($author['middleName']) : '';
            $lastName = isset($author['lastName']) ? $this->_sanitizeText($author['lastName']) : '';
            $name = $firstName;
            if ('' !== $middleName) {
                if ('' !== $name) {
                    $name .= self::NBSP;
                }
                $name .= $middleName;
            }
            if ('' !== $lastName) {
                if ('' !== $name) {
                    $name .= self::NBSP;
                }
                $name .= $lastName;
            }
            $run->addText($name, self::FS_ARTICLE_AUTHOR);
            // affiliation number
            if (1 < count($abstract['authors']) && 1 < count($abstract['affiliations'])) {
                $super = self::NBSP;
                $affiliations = $author['affiliations'];
                sort($affiliations);
                foreach ($affiliations as $affiliationIdx => $affiliationNum) {
                    if (0 < $affiliationIdx) {
                        $super .= ','.self::NBSP;
                    }
                    $super .= $affiliationNum + 1;
                }
                $run->addText($super, self::FS_ARTICLE_AUTHOR_SUPER);
            }
        }
        // Affilication
        $run = $section->addTextRun(self::PS_ARTICLE_AFFILIATION);
        foreach ($abstract['affiliations'] as $affiliationNum => $affiliation) {
            if ($affiliationNum !== $affiliation['position']) {
                exit('Fatal Error: "affiliations:position" is not sorted (uuid:'.$affiliation['uuid'].')'.PHP_EOL);
            }
            if (0 < $affiliationNum) {
                // $run->addText(', ', self::FS_ARTICLE_AFFILIATION);
                $run->addTextBreak();
            }
            // affiliation number
            if (1 < count($abstract['authors']) && 1 < count($abstract['affiliations'])) {
                $super = ($affiliation['position'] + 1).self::NBSP;
                $run->addText($super, self::FS_ARTICLE_AFFILIATION_SUPER);
            }
            // name
            $department = isset($affiliation['department']) ? $this->_sanitizeText($affiliation['department']) : '';
            $section_ = isset($affiliation['section']) ? $this->_sanitizeText($affiliation['section']) : '';
            $address = isset($affiliation['address']) ? $this->_sanitizeText($affiliation['address']) : '';
            $country = isset($affiliation['country']) ? $this->_sanitizeText($affiliation['country']) : '';
            $name = $department;
            if ('' !== $section_) {
                if ('' !== $name) {
                    $name .= ', ';
                }
                $name .= $section_;
            }
            if ('' !== $address) {
                if ('' !== $name) {
                    $name .= ', ';
                }
                $name .= $address;
            }
            if ('' !== $country) {
                if ('' !== $name) {
                    $name .= ', ';
                }
                $name .= $country;
            }
            $run->addText($name, self::FS_ARTICLE_AFFILIATION);
        }
        // Email
        $email = $this->_sanitizeText($abstract['authors'][0]['mail']);
        $section->addText($email, self::FS_ARTICLE_EMAIL, self::PS_ARTICLE_EMAIL);
        // DOI
        $doi = isset($abstract['doi']) ? $this->_sanitizeText($abstract['doi']) : '';
        if ('' !== $doi) {
            $section->addText(self::HL_DOI.':'.$doi, self::FS_ARTICLE_DOI, self::PS_ARTICLE_DOI);
        }
        // Abstract
        $run = $section->addTextRun(self::PS_ARTICLE_ABSTRACT);
        $textArr = explode("\n", $abstract['text']);
        foreach ($textArr as $textIdx => $text) {
            $text = $this->_sanitizeText($text);
            if ('' === $text) {
                continue;
            }
            if (0 < $textIdx) {
                $run = $section->addTextRun(self::PS_ARTICLE_ABSTRACT_FOLLOW);
            }
            $this->_addRunTextWithMath($run, $text, self::FS_ARTICLE_ABSTRACT);
        }
        // Figure
        foreach ($abstract['figures'] as $figure) {
            // Image
            $this->_printInfo(' >> Figure: '.$figure['uuid']);
            $imagePath = $this->mImagesDir.'/'.$figure['uuid'];
            if (!file_exists($imagePath)) {
                exit('Fatal Error: "figures" image file not found (uuid:'.$figure['uuid'].')'.PHP_EOL);
            }
            $run = $section->addTextRun(self::PS_ARTICLE_FIGURE);
            $imagePath2 = dirname($imagePath).'/'.basename($imagePath, '.png').'_w.png';
            if (false === ImageConverter::convertToWhiteBackgroundPNG($imagePath, $imagePath2)) {
                exit('Fatal Error: failed to create white backgroup image file from ('.$imagePath.')'.PHP_EOL);
            }
            list($width, $height) = $this->_getPreferedImageSize($imagePath2);
            $run->addImage($imagePath2, [
                'width' => PhpWord\Shared\Converter::cmToPoint($width / 10.0),
                'height' => PhpWord\Shared\Converter::cmToPoint($height / 10.0),
            ]);
            // Caption
            $caption = isset($figure['caption']) ? $this->_sanitizeText($figure['caption']) : '';
            if ('' !== $caption) {
                $run = $section->addTextRun(self::PS_ARTICLE_FIGURE_CAPTION);
                $heading = self::HL_FIGURE_CAPTION;
                $heading .= $figure['position'];
                $heading .= ': ';
                $run->addText($heading, self::FS_ARTICLE_FIGURE_CAPTION_HEADING);
                $this->_addRunTextWithMath($run, $caption, self::FS_ARTICLE_FIGURE_CAPTION);
            }
        }
        // Ackknowledgements
        $acknowledgements = isset($abstract['acknowledgements']) ? $this->_sanitizeText($abstract['acknowledgements']) : '';
        if ('' !== $acknowledgements) {
            $section->addText(self::HL_ACKNOWLEDGEMENTS, self::FS_ARTICLE_ACKNOWLEDGEMENTS_HEADING, self::PS_ARTICLE_ACKNOWLEDGEMENTS_HEADING);
            $run = $section->addTextRun(self::PS_ARTICLE_ACKNOWLEDGEMENTS);
            $this->_addRunTextWithMath($run, $acknowledgements, self::FS_ARTICLE_ACKNOWLEDGEMENTS);
        }
        // References
        if (count($abstract['references'])) {
            $name = self::NS_ARTICLE_REFERENCES.$numArticles;
            $section->addText(self::HL_REFERENCES, self::FS_ARTICLE_REFERENCES_HEADING, self::PS_ARTICLE_REFERENCES_HEADING);
            foreach ($abstract['references'] as $reference) {
                $text = isset($reference['text']) ? $this->_sanitizeText($reference['text']) : '';
                $link = isset($reference['link']) ? $this->_sanitizeText($reference['link']) : '';
                $doi = isset($reference['doi']) ? $this->_sanitizeText($reference['doi']) : '';
                if (preg_match('/(10.\d+\/.+)/', $doi, $matches)) {
                    $doi = $matches[1];
                }
                if ('' !== $link) {
                    if ('' === $text) {
                        $text .= $link;
                    }
                }
                if ('' !== $doi && !preg_match('/'.preg_quote($doi, '/').'/', $text)) {
                    if ('' !== $text) {
                        $text .= ', ';
                    }
                    $text .= self::HL_DOI.':'.$doi;
                }
                $section->addListItem($text, 0, self::FS_ARTICLE_REFERENCES, $name, self::PS_ARTICLE_REFERENCES);
            }
        }
        // Topic
        if (false) {
            $run = $section->addTextRun(self::PS_ARTICLE_TOPIC);
            $heading = self::HL_TOPIC.': ';
            $run->addText($heading, self::FS_ARTICLE_TOPIC_HEADING);
            $topics = $abstract['topic'];
            $run->addText($topics, self::FS_ARTICLE_TOPIC);
        }
        // author short name
        $authorShortNames = [];
        foreach ($abstract['authors'] as $authorIdx => $author) {
            $firstName = isset($author['firstName']) ? trim($this->_sanitizeText($author['firstName']), '*') : '';
            $middleName = isset($author['middleName']) ? $this->_sanitizeText($author['middleName']) : '';
            $lastName = isset($author['lastName']) ? $this->_sanitizeText($author['lastName']) : '';
            $shortName = $lastName;
            if ('' !== $firstName) {
                if ('' !== $shortName) {
                    $shortName .= ' ';
                }
                $shortName .= strtoupper(substr($firstName, 0, 1));
            }
            if ('' !== $middleName) {
                $middleNames = explode(' ', $middleName);
                foreach ($middleNames as $m) {
                    $shortName .= strtoupper(substr($m, 0, 1));
                }
            }
            $authorShortNames[] = $shortName;
        }
        // Citation
        $run = $section->addTextRun(self::PS_ARTICLE_CITATION);
        $heading = self::HL_CITATION.': ';
        $run->addText($heading, self::FS_ARTICLE_CITATION_HEADING);
        $text = implode(', ', $authorShortNames);
        $text .= ' ('.$this->_getPublicationYear().') ';
        $text .= $this->_sanitizeText($abstract['title']).'. ';
        if (true) {
            $text .= $this->_sanitizeText($this->mConference['name']);
            $text .= '. ';
            if (isset($abstract['doi'])) {
                $text .= self::HL_DOI.':'.$this->_sanitizeText($abstract['doi']);
            }
        } else {
            // aini2016 style.
            $text .= $this->_sanitizeText($this->_getBookTitle()).'. ';
            $text .= $this->_sanitizeText($this->_getConferenceTitle()).' ';
            $text .= $this->_sanitizeText($this->_getBookSubTitle()).': ';
            $text .= $this->_sanitizeText($this->_getSessionTitle($abstract['abstrTypes'][0]['short'])).'-';
            $text .= ($abstract['sortId'] & 0xFFFF).'. ';
            $text .= self::HL_DOI.':'.$this->_sanitizeText($abstract['doi']);
        }
        $run->addText($text, self::FS_ARTICLE_CITATION);
        // Copyright
        $run = $section->addTextRun(self::PS_ARTICLE_COPYRIGHT);
        $heading = self::HL_COPYRIGHT.': ';
        $run->addText($heading, self::FS_ARTICLE_COPYRIGHT_HEADING);
        $text = self::COPYRIGHT_SIGN.' ('.$this->_getPublicationYear().') ';
        $text .= implode(', ', $authorShortNames);
        $run->addText($text);
    }

    /**
     * add colophon page.
     *
     * @param \PhpOffice\PhpWord\Element\Section $section
     */
    protected function _addColophonPage($section)
    {
        // book title
        $textArr = explode("\n", $this->_getBookTitle());
        $run = $section->addTextRun(self::PS_COLOPHON_BOOK_TITLE);
        foreach ($textArr as $num => $text) {
            if (0 < $num) {
                $run->addTextBreak();
            }
            $run->addText($this->_sanitizeText($text), self::FS_COLOPHON_BOOK_TITLE);
        }
        // conference title
        $textArr = explode("\n", $this->_getConferenceTitle());
        $run = $section->addTextRun(self::PS_COLOPHON_CONFERENCE_TITLE);
        foreach ($textArr as $num => $text) {
            if (0 < $num) {
                $run->addTextBreak();
            }
            $run->addText($this->_sanitizeText($text), self::FS_COLOPHON_CONFERENCE_TITLE);
        }
        $run->addText(' '.self::HL_ABSTRACTS, self::FS_COLOPHON_CONFERENCE_TITLE);
        // publication date
        $text = trim(implode(' ', [$this->_getPublicationMonth(), $this->_getPublicationYear()]));
        $section->addText($this->_sanitizeText($text), self::FS_COLOPHON_PUBLICATION_DATE, self::PS_COLOPHON_PUBLICATION_DATE);
        // editor
        $section->addText(self::HL_EDITED_BY, self::FS_COLOPHON_EDITOR_HEADING, self::PS_COLOPHON_EDITOR_HEADING);
        $text = '';
        $editors = $this->_getEditor();
        foreach ($editors as $num => $editor) {
            if ($num > 0) {
                if ($num == (count($editors) - 1)) {
                    $text .= ' and ';
                } else {
                    $text .= ', ';
                }
            }
            $text .= $editor;
            ++$num;
        }
        $section->addText($this->_sanitizeText($text), self::FS_COLOPHON_EDITOR, self::PS_COLOPHON_EDITOR);
        // publisher
        $section->addText(self::HL_PUBLISHED_BY, self::FS_COLOPHON_PUBLISHER_HEADING, self::PS_COLOPHON_PUBLISHER_HEADING);
        $run = $section->addTextRun(self::PS_COLOPHON_PUBLISHER);
        $run->addText($this->_sanitizeText($this->_getPublisherName()), self::FS_COLOPHON_PUBLISHER);
        $run->addTextBreak();
        $run->addText($this->_sanitizeText($this->_getPublisherAddress()), self::FS_COLOPHON_PUBLISHER);
        // copyright
        $text = self::COPYRIGHT_SIGN.' ('.$this->_getPublicationYear().') '.$this->_getPublisherName();
        $section->addText($this->_sanitizeText($text), self::FS_COLOPHON_COPYRIGHT, self::PS_COLOPHON_COPYRIGHT);
    }

    /**
     * add run text with math.
     *
     * @param $run
     * @param string $text
     * @param string $style
     */
    protected function _addRunTextWithMath($run, $text, $style)
    {
        $strArr = preg_split('/(\$.+\$)/Uu', $text, -1, PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE);
        foreach ($strArr as $str) {
            $isNewLine = false;
            $mathStr = '';
            if (preg_match('/^\$(.+)\$$/', $str, $matches)) {
                $mathStr = $matches[1];
                if (preg_match('/^\$(.*)\$$/', $mathStr, $matches)) {
                    $mathStr = $matches[1];
                    $isNewLine = true;
                }
            }
            if ('' !== $mathStr) {
                $mathStr = preg_replace_callback('/textbf{(.*)}/Uu', function ($m) { return 'textbf{'.str_replace(' ', '\ ', $m[1]).'}'; }, $mathStr);
                $this->_printInfo(' >> Math: '.$mathStr);
                $imagePath = $this->mTexvc->createImage($mathStr);
                if ('' === $imagePath) {
                    exit('Fatal Error: failed to create tex image file from string ('.$mathStr.')'.PHP_EOL);
                }
                $imagePath2 = dirname($imagePath).'/'.basename($imagePath, '.png').'_w.png';
                if (false === ImageConverter::convertToWhiteBackgroundPNG($imagePath, $imagePath2)) {
                    exit('Fatal Error: failed to create white backgroup image file from ('.$imagePath.')'.PHP_EOL);
                }
                $fontStyle = PhpWord\Style::getStyle($style);
                $size = $fontStyle->getSize() ?: $this->mDefaultFontSize;
                $ratio = $size / self::TEXVC_FONT_SIZE;
                $imageSize = @getimagesize($imagePath2);
                list($width, $height) = $imageSize;
                $widthMm = $width / self::FIGURE_DPI * 25.4 * $ratio;
                $heightMm = $height / self::FIGURE_DPI * 25.4 * $ratio;
                if ($isNewLine) {
                    $run->addTextBreak();
                }
                $img = $run->addImage($imagePath2, [
                    'width' => PhpWord\Shared\Converter::cmToPoint($widthMm / 10.0),
                    'height' => PhpWord\Shared\Converter::cmToPoint($heightMm / 10.0),
                ]);
                if ($isNewLine) {
                    $run->addTextBreak();
                }
            } else {
                $run->addText($str, $style);
            }
        }
    }

    /**
     * get prefered image path.
     *
     * @param string $imagePath
     *
     * @return array
     */
    protected function _getPreferedImageSize($imagePath)
    {
        $size = @getimagesize($imagePath);
        if (false === $size) {
            return [0, 0];
        }
        list($width, $height) = $size;
        $widthMm = $width / self::FIGURE_DPI * 25.4;
        $heightMm = $height / self::FIGURE_DPI * 25.4;
        if ($widthMm > $this->mMaxFigureWidth) {
            $heightMm = $heightMm * $this->mMaxFigureWidth / $widthMm;
            $widthMm = $this->mMaxFigureWidth;
        }
        if ($heightMm > $this->mMaxFigureHeight) {
            $widthMm = $widthMm * $this->mMaxFigureHeight / $heightMm;
            $heightMm = $this->mMaxFigureHeight;
        }
        if ($widthMm > $this->mMaxFigureWidth * 0.8 && $heightMm > $this->mMaxFigureHeight * 0.8) {
            $widthMm *= 0.8;
            $heightMm *= 0.8;
        }
        $widthMm = intval($widthMm);
        $heightMm = intval($heightMm);

        return [$widthMm, $heightMm];
    }

    /**
     * sanitize text.
     *
     * @param string $text
     *
     * @return string
     */
    protected function _sanitizeText($text)
    {
        $text = str_replace(
            ["\t", '&', '<', '>', '"', '\''],
            [' ', '&amp;', '&lt;', '&gt;', '&quot;', '&apos;'],
            $text
        );
        $text = preg_replace('/[\x00-\x09\x0b\x0c\x0e-\x1f\x7f]/u', '', $text);
        $text = preg_replace('/ +/', ' ', $text);
        $text = trim($text);

        return $text;
    }

    /**
     * print information.
     *
     * @param string $text
     */
    protected function _printInfo($text)
    {
        echo '[INFO] '.$text.PHP_EOL;
    }
}
