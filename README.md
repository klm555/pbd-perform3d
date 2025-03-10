PBD_p3d
=========
![Static Badge](https://img.shields.io/badge/python-3.9.12-%233776AB?style=plastic&logo=Python)
![Static Badge](https://img.shields.io/badge/PyQt-5.15.7-%2341CD52?style=plastic&logo=Qt)
![Static Badge](https://img.shields.io/badge/Sphinx-5.3.0-%23000000?style=plastic&logo=Sphinx)

**PBD_p3d** 는 Perform-3D를 이용한 성능기반 내진설계를 보다 쉽고 빠르게 진행하기 위해 만들어진 Library입니다.

# Quick Start
## Clone
1. Github의 Repository를 clone할 로컬 디렉토리 지정. (예시)
```cmd
cd /c/user/Desktop
```
2. clone
```cmd
git clone https://github.com/klm555/structural-drawing-review.git
```
> [!WARNING]
> git bash가 설치되어있지 않은 경우, clone이 실행되지 않을 수 있음. git bash 설치 방법은 인터넷 참조. git bash를 설치하지 않고 clone을 진행하려면, github의 [레포지토리](https://github.com/klm555/pbd-perform3d)에서 초록 버튼을 누르고 다운로드하면 됨.

## Set Virtual Environment
3. Clone된 Repository의 디렉토리로 이동. (예시)
```cmd
cd .../user/Desktop/pbd-perform3d
```
4. [Python 3.9.12](https://www.python.org/downloads/release/python-3912/) 다운로드 후 설치
5. 아래의 명령어 차례대로 입력
```cmd
virtualenv .venv --python==python3.9.12
pip install -r requirements.txt
```
> [!WARNING]
> virtualenv가 설치되어있지 않은 경우, 두번쨰 명령어가 실행되지 않을 수 있음. 이 경우 `pip install virtualenv`로 virtualenv 라이브러리를 먼저 설치한 후 실행하면 됨.

6. 커맨드 프롬프트(cmd)에서  가상환경 활성화
```cmd
cd .venv/Scripts
activate
cd ../..
```

# Build `.exe` File
1. 디렉토리가 제대로 설정되어 있는지 확인. (예시)
```cmd
 cd .../user/Desktop/pbd-perform3d
```
2. `setup.py` 파일 생성 (현재는 이미 있으므로 따로 생성할 필요 없음)
3. 가상환경 활성화
```cmd
cd .venv/Scripts
activate
cd ../..
```
4. Build
```cmd
python setup.py build
```