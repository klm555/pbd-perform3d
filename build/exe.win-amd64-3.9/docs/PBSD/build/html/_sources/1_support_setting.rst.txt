==============================
지점 및 다이어프램 설정
==============================

지점 설정
^^^^^^^^^^^^
성능기반 내진설계에서는 원칙적으로 기초면 하부가 고정된 모델을 사용하여야 합니다. [#]_
그러나 경우에 따라 지하구조 측면의 구속효과를 고려하거나 지표면에서 고정조건을 사용하는 경우도 있으므로, 설계자와의 협의 이 후에 결정합니다.

본 성능기반 내진설계 업무절차서에서는 기초면 하부가 고정된 경우에 대해 모델링합니다.

.. topic:: What to do

   1. 기초면 하부에 지점을 설정하기에 앞서, 탄성설계 단계에서 설정된 지점과의 혼동을 피하기 위해 탄성설계 단계에서 설정된 지점을 삭제합니다. 
      
      :guilabel:`Tree Menu`\의 :guilabel:`Works` 탭에서 :guilabel:`Boundaries` - :guilabel:`Supports`\를 선택합니다.

      .. only:: html

         .. figure:: _static/images/2_delete_existing_supports.gif
            :align: center
            
   2. 설정되어있는 Supports를 마우스 오른쪽 버튼으로 클릭한 후 :kbd:`Delete`\을 눌러 모두 삭제합니다.
   
   3. 기초면 하부를 고정단으로 설정하기 위하여 Midas Gen에서 :guilabel:`Boundary` - :guilabel:`Define Supports`\를 클릭합니다.

   4. 모델의 가장 하단(기초면)을 선택합니다.

      .. figure:: _static/images/1_기초면_선택.png
         :align: center
         :scale: 90%

   5. Dx, Dy, Dz를 체크하고 :kbd:`Apply`\를 클릭합니다.

      설정이 완료되면, 기초면에 생성한 지점이 표시됩니다.

      .. figure:: _static/images/1_supports_설정.png
         :align: center
         :scale: 90%
   
다이어프램 설정
^^^^^^^^^^^^^^^^^^^
비선형해석에서 슬래브의 모델링은 전체 층에 대하여 면내강체로 정의되어야합니다. [#]_
따라서 전체 층이 면내강체(Rigid Diaphragm)로 설정되어있는지 확인하고, 면내강체가 설정되어있지 않은 층에는 면내강체를 설정합니다.
다만 지점조건이나 Ground Level에 따라 다이어프램의 설정 역시 달라질 수 있으므로, 설계자와 협의하여 결정합니다.

.. topic:: What to do

   1. Rigid Diaphragm 설정을 위하여 :guilabel:`Structure`\ - :guilabel:`Control Data...` - :guilabel:`Story...`\을 클릭합니다.

   2. 생성된 창에서 가장 아래층(기초면)을 제외한 모든 층의 ``Floor Diaphragm``\이 Consider로 설정되어있는지 확인합니다.

      .. figure:: _static/images/1_diaphragm_설정.png
         :align: center
         :scale: 90%
     
   3. Consider로 설정되어있지 않은 층(Do not Consider로 설정된 층)은 Consider로 변경합니다. 설정 완료 후 :kbd:`Close`\를 클릭합니다.

.. [#] 한국지진공학회, 철근콘크리트 건축물 성능기반 내진설계 지침 및 모델링 가이드(2019) 3.4-(1)
.. [#] 대한건축학회, 철근콘크리트 건축구조물의 성능기반 내진설계 지침(2021), 4.5-(1)