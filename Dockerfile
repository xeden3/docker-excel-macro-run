FROM xeden3/docker-office-python-core:v1
# FROM akkuman/msoffice2010-python:latest
MAINTAINER JamesChan "JamesChan<james@sctmes.com> (http://www.sctmes.com)"

RUN xvfb-run wine pip install pywin32

RUN apt-get update && apt-get install -y locales
RUN sed -i -e 's/# zh_CN.UTF-8 UTF-8/zh_CN.UTF-8 UTF-8/' /etc/locale.gen && \
    dpkg-reconfigure --frontend=noninteractive locales && \
    update-locale LANG=zh_CN.UTF-8
ENV LC_ALL=zh_CN.UTF-8

# 将工作目录设置为/opt/wineprefix/drive_c/
WORKDIR /opt/wineprefix/drive_c/

COPY libs/tini /tini
COPY code/excel_xlsm_macro_run.py /opt/wineprefix/drive_c/app/
COPY code/entrypoint.sh /entrypoint.sh

RUN chmod +x /tini && \
    chmod +x /entrypoint.sh

# 设置ENTRYPOINT
# ENTRYPOINT ["/tini", "--", "xvfb-run", "-a", "wine", "python", "/opt/wineprefix/drive_c/app/excel_xlsm_macro_run.py"]
ENTRYPOINT ["/tini", "--", "/entrypoint.sh"]
