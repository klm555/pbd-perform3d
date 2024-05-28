==========================================================
탄성설계 모델 Import
==========================================================

성능설계 모델은 Perform-3D로 직접 모델링할 수 있지만, 
빠르고 편리한 모델링을 위하여 레퍼런스 모델과 Data Conversion Sheets의 정보를 Import하여 모델링할 수 있습니다.

파일 변환 (csv 파일 생성)
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
첫 번째 절차는 Midas Gen 모델의 정보를 Perform-3D에서 읽어들일 수 있는 방식으로 변환하는 것입니다.
변환이 가능한 정보는 :doc:`Data Conversion Sheets 작성 <1_material_setting>` 장에서 모두 입력되었으므로, 
Data Conversion Sheets를 Perform-3D에서 읽을 수 있는 파일 형식인 ``.csv``\으로 변환합니다.

.. topic:: What to do

   1. PBD_p3d를 실행합니다.

   2. Data Conversion (Excel Sheets)에 Data Conversion Sheets의 경로를 입력합니다.

   3. Import에서 Import를 원하는 항목들을 체크합니다. 
      Nodal Loads를 Import하는 경우, Dead Load Name과 Live Load Name에 각각 Midas Gen에서 사용하였던 고정하중, 활하중 이름을 입력합니다.      

      .. figure:: _static/images/3_import_csv_설정.png
         :align: center
         :scale: 80%

   4. Import를 클릭합니다. Import가 완료되면 아래의 상태창에 ``Completed!``\가 표시됩니다. 
      또한 Data Conversion Sheets가 위치하는 경로에 아래와 같이 선택한 항목들의 ``.csv`` 형식 파일이 생성됩니다.

      .. figure:: _static/images/3_import_csv_생성.png
         :align: center
         :scale: 80%

Perform-3D 실행
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Perform-3D로 새 성능설계 모델 파일을 만들고, 앞서 생성한 csv 파일을 Import합니다. ``Materials``
새 성능설계 모델 파일은 새 파일을 생성하여 성능설계 모델의 모델링을 시작할 수도 있지만, 

.. note::
   Perform-3D는 Midas Gen과 같이 단일 파일( ``.mgb`` )로 생성 또는 저장되는 것이 아니라, 폴더 형태로 생성 또는 저장됩니다.
   따라서 Perform-3D 파일을 Load하는 경우, 파일이 아닌 폴더를 선택하여 Load합니다.
   

.. topic:: What to do

   1. Perform-3D를 실행한 후, :kbd:`Start a New Structure`\를 클릭합니다.

      .. figure:: _static/images/3_p3d_실행.png
         :align: center
         :scale: 90%

      .. raw:: latex
         
         \newpage 

   2. Structure Name(파일명), Location of STRUCTURES folder(파일 경로), Structure Description(파일 설명) 등을 입력합니다.
      파일명은 **영문 또는 숫자로, 띄어쓰기 없이** 입력해야 합니다.

      .. figure:: _static/images/3_p3d_실행_2.png
         :align: center
         :scale: 80%

      단위는 :doc:`단위 설정 <1_unit_setting>` 장에 따라 :math:`kN, mm`\로 설정합니다.

      Minimum spacing between nodes는 :math:`50 mm`\로 설정합니다.

      설정이 완료되면 :kbd:`OK`\를 클릭합니다.

.. raw:: latex
   
   \newpage

Import Nodes
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. topic:: What to do
   
   1. Nodes를 Import하기 위해 :guilabel:`Import/Export Structure Data`\을 클릭하고, 생성된 창에서 :kbd:`Import`\를 클릭합니다.

      .. only:: html

         .. figure:: _static/images/3_import_nodes.gif
            :align: center

   2. :guilabel:`Nodes Only` 탭을 클릭한 뒤, Specify name of text file에 앞서 생성한 ``Node.csv`` 파일의 경로를 입력합니다.


      .. figure:: _static/images/3_import_nodes_설정.png
         :align: center
         :scale: 90%

      .. raw:: latex

         \newpage
      
   3. 입력 후 :kbd:`Test` - :kbd:`확인`\을 클릭하면 아래와 같이 Import할 Nodes가 표시됩니다.
      
      .. figure:: _static/images/3_import_nodes_완료.png
         :align: center
         :scale: 80%

      :kbd:`OK`\를 클릭하여 Import를 완료합니다.

Import Masses
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. topic:: What to do

   1. :guilabel:`Nodes`\를 클릭하고 생성된 창에서 :guilabel:`Masses` 탭을 클릭합니다. 
      Mass Pattern을 만들기 위해 :kbd:`New`\를 클릭합니다. 
   
      .. only:: html

         .. figure:: _static/images/3_import_masses.gif
            :align: center

   2. Enter pattern name에 ``Mass``\(또는 사용자가 원하는 이름)를 입력한 후, :kbd:`OK`\를 눌러 Mass Patter 생성을 완료합니다.
   
      .. figure:: _static/images/3_import_masses_패턴생성.png
         :align: center
         :scale: 90%
   
   3. Masses를 Import하기 위해 :guilabel:`Import/Export Structure Data`\을 클릭하고, 생성된 창에서 :kbd:`Import`\를 클릭합니다.

   4. :guilabel:`Masses` 탭을 클릭한 뒤, Choose mass pattern에서 방금 생성한(또는 사용자가 원하는) Mass Pattern을 선택합니다.
      Specify name of text file에 앞서 생성한 ``Mass.csv`` 파일의 경로를 입력합니다.

      .. raw:: latex

         \newpage

      .. figure:: _static/images/3_import_masses_설정.png
         :align: center
         :scale: 90%

      .. raw:: latex

         \newpage

   5. 입력 후 :kbd:`Test` - :kbd:`확인`\을 클릭하면 아래와 같이 Import할 Nodes가 표시됩니다.

      .. figure:: _static/images/3_import_masses_완료.png
         :align: center
         :scale: 80%

      :kbd:`OK`\를 클릭하여 Import를 완료합니다.

.. raw:: latex

   \newpage

Import Elements
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. topic:: What to do

   1. :guilabel:`Elements`\를 클릭하고, Elements Group을 만들기 위해 생성된 창에서 :kbd:`New`\를 클릭합니다. 
   
      .. only:: html

         .. figure:: _static/images/3_import_elements.gif
            :align: center

   2. 먼저 연결보를 Import하기 위해 연결보(Coupling Beam) 그룹을 생성합니다.
      Element Type에서 ``Beam``\을 선택한 후,
      Group Name에 ``C.Beam``\(또는 사용자가 원하는 이름)을 입력합니다.
      
      .. figure:: _static/images/3_import_elements_그룹생성.png
         :align: center
         :scale: 90%

      :kbd:`OK`\를 눌러 연결보 그룹 생성을 완료합니다.

   3. 같은 방법으로 연결보 외의 Import할 Elements와 Gages 그룹도 생성합니다.

      :연결보: Element Type: ``Beam`` / Group Name: ``C.Beam``
      :일반보: Element Type: ``Beam`` / Group Name: ``G.Beam``
      :탄성보: Element Type: ``Beam`` / Group Name: ``E.Beam``
      :일반기둥: Element Type: ``Column`` / Group Name: ``G.Column``
      :탄성기둥: Element Type: ``Column`` / Group Name: ``E.Column``
      :벽체: Element Type: ``Shear Wall`` / Group Name: ``Wall``
      :벽체 회전각 게이지: Element Type: ``Deformation Gage`` / Group Name: ``Wall Rotation`` / Gage Type: ``Wall type, rotation or shear``
      :벽체 축변형률 게이지: Element Type: ``Deformation Gage`` / Group Name: ``Wall Axial Strain`` / Gage Type: ``Bar type, axial strain``
      :기둥 회전각 게이지(X): Element Type: ``Deformation Gage`` / Group Name: ``Column Rotation(X)`` / Gage Type: ``Beam type, rotation``
      :기둥 회전각 게이지(Y): Element Type: ``Deformation Gage`` / Group Name: ``Column Rotation(Y)`` / Gage Type: ``Beam type, rotation``

   4. Import하지 않지만, 이 후의 모델링 과정에서 사용자가 직접 생성할 Elements와 Gages의 그룹도 같은 방법으로 생성할 수 있습니다.

      :Imbedded Beam: Element Type: ``Beam`` / Group Name: ``Imbedded Beam``
      :보 회전각 게이지: Element Type: ``Deformation Gage`` / Group Name: ``Beam Rotation`` / Gage Type: ``Beam type, rotation``
   
   5. Masses를 Import하기 위해 :guilabel:`Import/Export Structure Data`\을 클릭하고, 생성된 창에서 :kbd:`Import`\를 클릭합니다.

   6. :guilabel:`Masses` 탭을 클릭한 뒤, Choose mass pattern에서 방금 생성한(또는 사용자가 원하는) Mass Pattern을 선택합니다.
      Specify name of text file에 앞서 생성한 ``Mass.csv`` 파일의 경로를 입력합니다.

      .. raw:: latex

         \newpage

      .. figure:: _static/images/3_import_masses_설정.png
         :align: center
         :scale: 90%

      .. raw:: latex

         \newpage

   7. 입력 후 :kbd:`Test` - :kbd:`확인`\을 클릭하면 아래와 같이 Import할 Nodes가 표시됩니다.

      .. figure:: _static/images/3_import_masses_완료.png
         :align: center
         :scale: 80%

      :kbd:`OK`\를 클릭하여 Import를 완료합니다.