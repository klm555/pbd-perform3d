PBD_p3d
=========
![Static Badge](https://img.shields.io/badge/python-3.9.12-%233776AB?style=plastic&logo=Python)
![Static Badge](https://img.shields.io/badge/PyQt-5.15.7-%2341CD52?style=plastic&logo=Qt)
![Static Badge](https://img.shields.io/badge/Sphinx-5.3.0-%23000000?style=plastic&logo=Sphinx)

**PBD_p3d** 는 Perform-3D를 이용한 성능기반 내진설계를 보다 쉽고 빠르게 진행하기 위해 만들어진 Library입니다.


# Build `.exe` File
1. 디렉토리가 제대로 설정되어 있는지 확인. (예시)
```cmd
 cd .../user/pbd-perform3d
```
2. `setup.py` 파일 생성 (현재는 이미 있으므로 따로 생성할 필요 없음)
3. 커맨드 프롬프트(cmd)에서 가상환경 활성화
```cmd
cd .venv/Scripts
activate
cd ../..
```
4. Build
```cmd
python setup.py build
```