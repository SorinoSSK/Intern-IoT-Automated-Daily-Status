# How To Use

## Requirements
1. Ensure you have [Python](https://www.python.org/downloads/) downloaded
2. Open CLI 
3. At the root, do:
    ```
    pip install -r requirements.txt
    ```
    Or
    ```
    pip3 install -r requirements.txt
    ```
    depending on your version of Python.

## Run Script
1. Open CLI of choice / launch `unit_status.py` with Python's IDLE.
2. Run script via IDLE or on the CLI with 
    ```
    python unit_status.py
    ```
    Or
    ```
    python3 unit_status.py
    ```

For more details, see the documentation under `documentation/`.

## Preparing Advantech board for python
1. Install linux library to add new repository
```sh
sudo apt install software-properties-common
```

2. Add python 3.11 repository
```sh
sudo add-apt-repository ppa:deadsnakes/ppa
```

3. Install pip for python 3.11
```sh
sudo apt install python3.11-distuils 
```
```
curl -sS https://bootstrap.pypa.io/get-pip.py | python3.11
```

4. Updata pip and libraries
```
python3.11 -m pip install --upgrade pip setuptools wheel
```

5. Download developer tools for wheels
```
sudo apt-get install python3.11-dev
```
6. You device is now ready to run [requirements](#requirements)


/lib/systemd/system/dailyStatus.service
systemctl status dailyStatus
systemctl start dailyStatus
systemctl restart dailyStatus
systemctl stop dailyStatus
systemctl enable dailyStatus
