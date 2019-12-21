FROM centos:7

ENV LANG C.UTF-8
ENV TZ=Asia/Shanghai
# Install required packages and remove the apt packages cache when done.

WORKDIR /etc/yum.repos.d/
RUN yum -y install wget && \
    wget http://mirrors.163.com/.help/CentOS6-Base-163.repo && \
    mv CentOS6-Base-163.repo CentOS-Base.repo && \
    yum makecache && \
    ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone && \
    yum install -y \
    python3 \
    python3-devel \
    python3-setuptools \
    python3-pip \
    nginx \
    python36-dateutil

WORKDIR /opt/workspace/FasterRunner/
COPY start.sh .
COPY manage.py .
COPY uwsgi.ini .
COPY requirements.txt .

RUN  pip3 install -r ./requirements.txt -i \
     https://pypi.douban.com/simple \
     --default-timeout=100

RUN  mkdir ./logs
RUN  touch ./logs/worker.log
RUN  touch ./logs/beat.log

EXPOSE 5000

CMD bash ./start.sh