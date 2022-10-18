<?php

namespace NIJC;

/**
 * texvc command wrapper class.
 */
class TexvcCommand
{
    public const TEXVC = 'texvc';
    public const ENCODING = 'UTF-8';

    /**
     * texvc command.
     *
     * @var string
     */
    protected $mTexvc;

    /**
     * output images directory.
     *
     * @var string
     */
    protected $mImagesDir;

    /**
     * temporary directory.
     *
     * @var NIJC\TempDirectory
     */
    protected $mTempDir;

    /**
     * constructor.
     *
     * @param string $imageDir
     */
    public function __construct($imageDir)
    {
        $this->mTexvc = getenv('TEXVC');
        if (false === $this->mTexvc) {
            $this->mTexvc = self::TEXVC;
        }
        $this->mTempDir = new TempDirectory();
        $this->mImagesDir = $imageDir;
    }

    /**
     * create image.
     *
     * @param string $tex
     * @param string $encoding
     *
     * @return string
     */
    public function createImage($tex, $encoding = 'UTF-8')
    {
        static $okval = [
            '+', // ok, but not html or mathml
            'c', // ok, conservative html, no mathml
            'm', // ok, moderate html, no mathml
            'l', // ok, liberal html, no mathml
            'C', // ok, conservative html, with mathml
            'M', // ok, moderate html, with mathml
            'L', // ok, liberal html, with mathml
            'X', // ok, no html, with mathml
        ];
        $cmd = $this->mTexvc.' '.escapeshellarg($this->mTempDir->get()).' '.escapeshellarg($this->mImagesDir).' '.escapeshellarg($tex).' '.escapeshellarg($encoding);
        $contents = shell_exec($cmd);
        if (0 == strlen($contents)) {
            // unknown error
            return '';
        }
        $retval = substr($contents, 0, 1);
        if (!in_array($retval, $okval)) {
            // error occured
            return '';
        }
        $hash = substr($contents, 1, 32);
        if (!preg_match('/^[a-f0-9]{32}$/', $hash)) {
            // unknown error
            return '';
        }
        $imagePath = $this->mImagesDir.'/'.$hash.'.png';
        if (!file_exists($imagePath)) {
            // unknown error
            return '';
        }

        return $imagePath;
    }
}
