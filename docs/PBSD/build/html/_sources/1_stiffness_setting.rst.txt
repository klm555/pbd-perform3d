==============
유효강성 설정
==============

성능기반 내진설계에서의 유효강성은 아래의 표\ [#]_ 에 따릅니다.

.. figure:: _static/images/유효강성.png
   :align: center
   
   *비선형모델의 유효강성*

.. note::
   유효강성은 Midas Gen에서 Perform-3D로 Import되는 항목이 아니지만, 
   이 후 생성할 레퍼런스 모델, 성능설계 모델과의 비교 검증을 위해 사용됩니다.

위의 표를 참고하여 Midas Gen에서 유효강성을 변경합니다. 다만 전단강성 :math:`GA_W` 의 경우, 계산에 필요한 단면적 :math:`A_W` 가 
유효단면적( :math:`A_e` )이 아닌 전체단면적( :math:`A_g` )임에 주의해야 합니다.
Midas Gen에서는 유효단면적(:math:`A_e = \frac{5}{6}A_g` ; 모든 보의 단면적은 직사각형으로 가정함)을 자동으로 계산하여 사용하므로, 
역수인 :math:`\frac{6}{5}(\approx 1.2)` 를 곱하여 전체단면적을 만들어 사용합니다.

연결보
^^^^^^^

연결보의 유효강성은 아래의 절차에 따라 변경, 추가합니다.

.. topic:: What to do

   1. Midas Gen에서 :guilabel:`Properties` - :guilabel:`Scale Factor` - :guilabel:`Section Stiffness Scale Factor` 를 클릭합니다.
 
   ..
    .. figure:: _static/images/section_stiffness_scale_factor.gif
       :align: center

   2. :guilabel:`Section Stiffness Scale Factor` 창에서 변경, 추가할 연결보의 Section을 선택한 후, Scale Factor를 변경하여 줍니다. 
      휨강성은 :math:`0.3EI` 이므로, :math:`I_{yy}, I_{zz}` 에 각각 :math:`0.3` 을 입력합니다. 
   
      .. image:: _static/images/c_beam_bending_유효강성.png
         :align: center
         :scale: 60%

      입력 후, :kbd:`Add/Replace` 버튼을 누릅니다. 

   3. 전단강성은 :math:`0.04(\frac{l}{h})GA` 이므로, :math:`A_{sy}, I_{sz}` 의 값을 변경해야 합니다. 
      연결보의 길이( :math:`l` )와 깊이( :math:`h` )를 확인한 후, :math:`0.04(\frac{l}{h})` 를 계산합니다.
      위의 설명과 같이, :math:`1.2` 를 곱합니다.
   
      .. figure:: _static/images/c_beam_shear_유효강성.png
         :align: center
         :scale: 60%

      .. warning::
         Midas Gen 모델링 과정에서 짧은 벽을 생략하는 경우, 연결보의 길이가 길게 모델링되는 경우가 있습니다.
         따라서 **도면을 확인** 하여 정확한 연결보의 길이를 이용해 계산합니다.

   4. 모든 연결보의 유효강성을 변경, 추가한 후, :kbd:`Close` 버튼을 누릅니다.

보, 기둥
^^^^^^^^^^^

연결보와 동일한 방식으로 :guilabel:`Section Stiffness Scale Factor` 에서 유효강성을 변경, 추가합니다. 
기둥의 휨강성 역시, Data Conversion Sheet에서 자동으로 계산된 값을 사용할 수 있습니다.

벽체
^^^^^^^


.. [#] 대한건축학회, 철근콘크리트 건축구조물의 성능기반 내진설계 지침(2021) [표 6-1]
.. [#] :math:`=5/6`