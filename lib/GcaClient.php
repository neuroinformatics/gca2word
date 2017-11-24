<?php

namespace NIJC;

/**
 * GCA client class.
 */
class GcaClient
{
    /**
     * GCA url.
     *
     * @var string
     */
    protected $mUrl;

    /**
     * user credential.
     *
     * @var string
     */
    protected $mUname;

    /**
     * user password.
     *
     * @var string
     */
    protected $mPassword;

    /**
     * temporary directory.
     *
     * @var NIJC\TempDirectory
     */
    protected $mTempDirectory;

    /**
     * cookie file path.
     *
     * @var string
     */
    protected $mCookieFilePath;

    /**
     * is authenticated.
     *
     * @var bool
     */
    protected $mIsAuthenticated = false;

    /**
     * constructor.
     *
     * @param string $url
     * @param string $uname
     * @param string $pass
     */
    public function __construct($url, $uname, $pass)
    {
        if ('/' === substr($url, -1)) {
            $url = substr($url, 0, -1);
        }
        $this->mUrl = $url;
        $this->mUname = $uname;
        $this->mPassword = $pass;
        $this->mTempDirectory = new TempDirectory();
        $this->mCookieFilePath = $this->mTempDirectory->get().'/cookie.dat';
    }

    /**
     * destructor.
     */
    public function __destruct()
    {
        $this->logout();
    }

    /**
     * authenticate.
     *
     * @return bool
     */
    public function authenticate()
    {
        if ($this->mIsAuthenticated) {
            return true;
        }
        $params = [
            'identifier' => $this->mUname,
            'password' => $this->mPassword,
        ];
        $url = $this->mUrl.'/authenticate/credentials?'.http_build_query($params);
        $res = $this->_httpGet($url);
        if (false === $res) {
            return false;
        }
        $this->mIsAuthenticate = true;

        return true;
    }

    /**
     * logout.
     *
     * @return bool
     */
    public function logout()
    {
        if ($this->mIsAuthenticated) {
            $url = $this->mUrl.'/logout';
            $res = $this->_httpGet($url);
            if (false === $res) {
                return false;
            }
        }
        $this->mIsAuthenticated = false;

        return true;
    }

    /**
     * get conference json.
     *
     * @param string $conference
     *
     * @return string|bool
     */
    public function getConference($conference)
    {
        $this->authenticate();
        $url = $this->mUrl.'/api/conferences/'.$conference;
        $data = $this->_httpGet($url);
        if (false === $data) {
            return false;
        }

        return $data;
    }

    /**
     * get abstracts json.
     *
     * @param string $conference
     *
     * @return string|bool
     */
    public function getAbstracts($conference)
    {
        $this->authenticate();
        $url = $this->mUrl.'/api/conferences/'.$conference.'/allAbstracts';
        $data = $this->_httpGet($url);
        if (false === $data) {
            return false;
        }

        return $data;
    }

    /**
     * get images.
     *
     * @param string $json
     *
     * @return bool
     */
    public function saveImages($json, $imageDir)
    {
        $this->authenticate();
        $abstracts = json_decode($json, true);
        if (null === $abstracts) {
            return false;
        }
        foreach ($abstracts as $abstract) {
            foreach ($abstract['figures'] as $figure) {
                $url = $figure['URL'];
                $fpath = $imageDir.'/'.$figure['uuid'];
                if (false === $this->_httpGet($url, $fpath)) {
                    return false;
                }
            }
        }

        return true;
    }

    /**
     * http get.
     *
     * @param string $url
     * @param string $fpath
     *
     * @return string|bool
     */
    protected function _httpGet($url, $fpath = '')
    {
        if (!empty($fpath)) {
            $fp = fopen($fpath, 'w');
            if (false === $fp) {
                return false;
            }
        }
        $curl = curl_init();
        curl_setopt($curl, CURLOPT_URL, $url);
        curl_setopt($curl, CURLOPT_FOLLOWLOCATION, true);
        curl_setopt($curl, CURLOPT_AUTOREFERER, true);
        curl_setopt($curl, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($curl, CURLOPT_FAILONERROR, true);
        curl_setopt($curl, CURLOPT_TIMEOUT, 5);
        curl_setopt($curl, CURLOPT_COOKIEFILE, $this->mCookieFilePath);
        curl_setopt($curl, CURLOPT_COOKIEJAR, $this->mCookieFilePath);
        if (!empty($fpath)) {
            curl_setopt($curl, CURLOPT_FILE, $fp);
        }
        $ret = curl_exec($curl);
        $code = curl_getinfo($curl, CURLINFO_HTTP_CODE);
        curl_close($curl);
        if (!empty($fpath)) {
            fclose($fp);
            $ret = true;
        }
        if (200 !== $code) {
            return false;
        }

        return $ret;
    }
}
