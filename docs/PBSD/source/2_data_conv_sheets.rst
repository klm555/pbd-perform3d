===============================================================
Data Conversion Sheet 작성
===============================================================
:bdg-success:`Excel`

Data Conversion Sheet는 성능기반 내진설계에 필요한 모든 정보를 입력할 엑셀 파일입니다. 
본 업무절차서에서 소개할 성능기반 내진설계의 모든 과정은 이 엑셀 파일을 기반으로 하며, 
따라서 이 파일이 제대로 작성되어야 모델링에서부터 결과 확인까지 오류없이 진행될 수 있습니다.

시트 작성에 앞서, 각 시트의 구성을 아래에 간략하게 소개하겠습니다.

.. glossary::

  ETC
     철근과 콘크리트의 강도 정보를 입력할 시트

  Materials
     철근과 콘크리트의 재료 물성치 정보 시트. 
     참조용(입력 또는 변경 X)
  
  Nodes / Elements / Nodal Loads / Story Mass / Story Data
     Midas Gen에서 Import할 정보를 입력할 시트

  Beam Naming / Column Naming / Wall Naming
     Naming에 필요한 정보를 입력할 시트

  C. Beam Properties / G.Column Properties / Wall Properties
     연결보, 기둥, 벽체의 모든 정보를 입력할 시트

  Output_Naming
     앞에서 입력한 정보들을 바탕으로 이름이 출력되는 시트. 출력용(사용자 입력 X)

  Output_G.Beam Properties / Output_E.Beam Properties / Output_E.Column Properties
     일반보, 탄성보, 탄성기둥의 모든 정보를 입력할 시트

  Output_C.Beam Properties / Output_G.Column Properties / Output_Wall Properties
     앞에서 입력한 정보들을 바탕으로 정리된 연결보, 일반기둥, 벽체의 정보가 출력되는 시트

  Results_C.Beam / Results_E.Beam / Results_Wall / Results_E.Column(개별) / ?????
     해석결과를 바탕으로 연결보, 탄성보, 벽체, ???의 강도 검토 결과가 출력되는 시트. 출력용(사용자 입력 X)


.. note::
   레퍼런스 모델은 아래와 같은 이유때문에 중요합니다.

   #. Perform-3D로 Import하기 위한 최종 모델이기때문에, 이 모델이 정확하지 않으면 성능설계 모델도 부정확하게 모델링됩니다.
   #. 탄성설계 모델에서 성능설계 모델로 변환하는 과정에서 변경 사항이 많다면, 레퍼런스 모델이 비교 검증에 가장 중요한 역할을 하는 모델이 됩니다.(?)

탄성설계 모델과 레퍼런스 모델의 생성 과정은 아래의 단계를 거칩니다.

.. toctree::
   :maxdepth: 1
   :caption: STEPS

   2_naming_rules


.. card:: 
    
   최종 탄성설계 모델을 복사(또는 Save as)하여 새로운 탄성설계 모델을 생성한 후, 파일을 열어줍니다.

