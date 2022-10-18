FROM docker.io/library/centos:7
WORKDIR /
RUN yum install -y epel-release \
  && yum install -y yum-plugin-copr \
  && yum copr -y enable orrisroot/gca2word \
  && yum install -y https://rpms.remirepo.net/enterprise/remi-release-7.rpm \
  && yum-config-manager --enable remi-php81 \
  && yum install -y texvc git unzip php-cli composer \
  && git clone https://github.com/neuroinformatics/gca2word.git \
  && cd gca2word \
  && composer install
ENTRYPOINT ["/gca2word/gca2word.php"]
