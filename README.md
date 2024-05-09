# Excel Macro Runner Docker (docker-excel-macro-run)

This Docker container provides a solution to run Excel macros using Python win32 on a Wine and Office environment. It supports installation via GitHub and Docker Hub. Below are the steps for installation and usage.

### Installation

#### GitHub Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/xeden3/docker-excel-macro-run.git
    cd docker-excel-macro-run
    ```

2. Build the Docker image:

    ```bash
    docker build -t docker-excel-macro-run:v1 .
    ```

#### Docker Hub Installation

Pull the Docker image directly from Docker Hub:

```bash
docker pull xeden3/docker-excel-macro-run:v1
```

### Usage

Run the Docker container with the following command:

```bash
docker run -v ./example.xlsm:/opt/wineprefix/drive_c/test.xlsm --rm docker-excel-macro-run:v1 test.xlsm ThisWorkbook.WriteDataToSheet1
```

The parameters explained:

- `docker-excel-macro-run:v1`: Docker image name and tag.
- `test.xlsm`: Name of the Excel file to run the macro on. It should match the filename in the container's directory (`/opt/wineprefix/drive_c/test.xlsm`).
- `ThisWorkbook.WriteDataToSheet1`: Macro command to execute.

The output will be in JSON format:

```json
{"errcode": 0, "errmsg": ""}
```

You can retrieve and print the output using:

```bash
output=$(docker run -v ./example.xlsm:/opt/wineprefix/drive_c/test.xlsm --rm docker-excel-macro-run:v1 test.xlsm ThisWorkbook.WriteDataToSheet1)
echo $output
```

### Challenges

This Docker container addresses two main challenges:

1. **Handling Errors with Chinese Programs**: Chinese characters are not recognized in the default Wine environment. Therefore, when macros contain Chinese or other languages, errors may occur. Installing the appropriate font library and changing the `locales` value can resolve this issue.

    ![4ded46e3f2b401979a661154c2ef4c4](https://github.com/xeden3/docker-excel-macro-run/assets/38025067/53f29124-d5ad-4a47-837b-5badecbc7ab4)

    To address this in Docker, the following steps are taken:

    ```Dockerfile
    RUN apt-get update && apt-get install -y locales
    RUN sed -i -e 's/# zh_CN.UTF-8 UTF-8/zh_CN.UTF-8 UTF-8/' /etc/locale.gen && \
        dpkg-reconfigure --frontend=noninteractive locales && \
        update-locale LANG=zh_CN.UTF-8
    ENV LC_ALL=zh_CN.UTF-8
    ```

### Challenges (Continued)

This Docker container also addresses another significant challenge:

2. **Suppressing xvfb-run Output**: During program execution, `xvfb-run` may output the following warning:

    ```
    X connection to :100 broken (explicit kill or server shutdown).
    ```
    ![image](https://github.com/xeden3/docker-excel-macro-run/assets/38025067/bb7777d0-ed07-48a7-8bbc-cca8da14413e)

    To mitigate this, an `entrypoint.sh` file is added. This shell script filters out the warning message using the `grep` command. Failure to suppress this warning may result in numerous exception messages being returned.

    Below is the code snippet used to suppress the warning:

    ```bash
    # Disable the 'X connection to :100 broken (explicit kill or server shutdown).' warning
    xvfb-run -a wine python /opt/wineprefix/drive_c/app/excel_xlsm_macro_run.py "$@" | grep -v '100 broken'
    ```

    This command ensures that the warning message is filtered out from the output, providing a cleaner execution result.
    
