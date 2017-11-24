<?php

namespace NIJC;

/**
 * temporary directory class.
 */
class TempDirectory
{
    /**
     * temporary directory.
     *
     * @var string
     */
    protected $mTempDir = null;

    /**
     * constructor.
     *
     * @param string $imageDir
     */
    public function __construct()
    {
        $tempPath = tempnam(sys_get_temp_dir(), str_replace('\\', '-', __CLASS__));
        if (file_exists($tempPath)) {
            unlink($tempPath);
            if (mkdir($tempPath)) {
                $this->mTempDir = $tempPath;
            }
        }
    }

    /**
     * destructor.
     */
    public function __destruct()
    {
        if (null !== $this->mTempDir) {
            $iterator = new \RecursiveDirectoryIterator($this->mTempDir);
            foreach (new \RecursiveIteratorIterator($iterator, \RecursiveIteratorIterator::CHILD_FIRST) as $file) {
                $fname = $file->getFilename();
                if (in_array($fname, ['.', '..'])) {
                    continue;
                }
                $fpath = $file->getPathname();
                if ($file->isDir()) {
                    rmdir($fpath);
                } else {
                    unlink($fpath);
                }
            }
            rmdir($this->mTempDir);
        }
    }

    /**
     * get temporary directory.
     *
     * @return string
     */
    public function get()
    {
        return $this->mTempDir;
    }
}
