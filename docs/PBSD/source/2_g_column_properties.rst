======================
G.Column Properties
======================

G.Column Properties 시트에는 일반기둥(General Column)에 대한 정보를 입력합니다.

각 열마다 입력해야할 정보는 다음과 같습니다.

* ``Name``\, ``Story``\, ``Dimensions``\, ``Cover``
    C.Beam Properties 시트와 동일한 방식으로 입력합니다.

    .. figure:: _static/images/2_c_beam_예시.png
       :align: center

* ``Rebar``
    철근 정보를 입력합니다.

    ``내진상세 여부`` 열에는 해당 철근의 내진상세 여부를 입력 또는 선택합니다.

    ``Main``\과 ``Hoop``\의 첫번째 열에는 각각 주근과 후프근이 일반용 철근인지 내진용 철근인지 입력 또는 선택합니다.

    ``Main``\과 ``Hoop``\의 두번째 열에는 각각 주근과 후프근의 철근 종류를 입력 또는 선택합니다.

* ``Arrangement``
    주근과 후프근의 개수와 간격을 입력합니다.
    ``Layer1`` 열에는 아래의 그림을 참조하여 가장 바깥 레이어의 철근 개수와 행의 개수를 입력합니다.

    일반기둥의 배근이 2단으로 되어있는 경우, 안쪽 레이어의 철근 개수와 행의 개수를 ``Layer2`` 열에 입력합니다.

    .. figure:: _static/images/2_g_column_properties_section.png
       :align: center
       :scale: 60%    

    ``Hoop``\의 ``X``\, ``Y`` 열에는 각각  X-브레이싱의 개수와 각도를 각각 입력합니다.

    .. figure:: _static/images/2_g_column_properties_section_2.png
       :align: center
       :scale: 60%

   부재일람표에 기둥의 Hoop(Mid)와 Hoop(End) 정보가 모두 있는 경우, Hoop(End)의 정보를 입력합니다.

.. topic:: What to do
    
   위의 정보를 참조하여 연결보의 정보를 입력합니다.