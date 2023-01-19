Python 사용법
================

본 장에서는 PBD_p3d를 사용하기 위해 필요한 최소한의 Python 사용에 대한 지식을 소개합니다.

더 많이 Python을 사용하고 공부하신다면, PBD_p3d에 사용된 함수들을 이해하거나 더 많은 기능을 직접 만들어 사용하실 수 있을 것입니다!

어디까지나 반대로, Python보다 엑셀이나 다른 프로그래밍 언어가 편하신 분들도 있을것이라고 생각합니다. 

PBD_p3d를 이용하기 위해서는 Python을 설치해야 합니다.

Python 코드를 편집(edit)하고 컴파일(compile)할 수 있는 환경을 제공하는 프로그램을 흔히 IDE(Integrated Development Environment)라고 합니다.
PyCharm, Jupyter Notebook, Visual Studio Code와 같은 다양한 IDE들이 쓰입니다만, 이 페이지에서는 Spyder를 이용한 Python 사용법을 소개해드릴 예정입니다.

설치
^^^^^^^^^

Spyder만 따로 설치가 가능하나, Spyder가 포함된 `Anaconda <https://www.anaconda.com/products/distribution>`_ 라는 패키지를 설치하는 것이 좋다.
Anaconda는 코드를 실행시키는데 필요한 여러 패키지를 대부분 포함하고 있기때문에, 패키지를 따로 설치하는 수고를 덜 수 있다.
`Anaconda <https://www.anaconda.com/products/distribution>`_ 를 다운로드받아서 설치한다.

Anaconda 및 Spyder 실행
^^^^^^^^^^^^^^^^^^^^^^^^^^^

수동 패키지 설치
^^^^^^^^^^^^^^^^^^^

PBD_p3d에는 Anaconda에서 제공하지 않는 패키지도 사용되기 때문에, 이 패키지들은 직접 설치해야합니다.

아래의 방법에 따라 설치하면 됩니다.

1. Anaconda Prompt를 실행합니다.

.. image:: /_static/install_1.png

2. 아래의 명령어를 입력 후 실행합니다.

.. code-block:: shell

   conda install -c conda-forge python-docx

.. note::
   docx는 해석 결과데이터를 워드 파일에 저장해주는 패키지입니다.




PBD_p3d는 Perform-3D를 이용한 성능기반 내진설계를 보다 쉽고 빠르게 진행하기 위해 만들어진 Python 모듈입니다.



