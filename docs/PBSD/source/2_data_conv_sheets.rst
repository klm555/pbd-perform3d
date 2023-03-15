===============================================================
Data Conversion Sheets 작성
===============================================================

.. only:: html
    
   :bdg-success:`Excel` :bdg-danger:`Midas GEN`

Data Conversion Sheet는 성능기반 내진설계에 필요한 모든 정보를 입력할 엑셀 파일입니다.
본 업무절차서에서 소개할 성능기반 내진설계의 모든 과정은 이 엑셀 파일을 기반으로 하며, 
따라서 이 파일이 제대로 작성되어야 모델링에서부터 결과 확인까지 오류없이 진행될 수 있습니다.

시트 작성에 앞서, 각 시트의 구성을 간략하게 소개합니다.

* :doc:`ETC <2_etc>`
    철근과 콘크리트의 강도 정보를 입력할 시트.


* :doc:`Materials <2_materials>`
    철근과 콘크리트의 재료 물성치 정보 시트. 참조용.


* :doc:`Nodes / Elements <2_nodes_elements>` / :doc:`Nodal Loads <2_nodal_loads>` / :doc:`Story Mass / Elements <2_story_mass>` / Story Data
    Midas Gen에서 Import할 정보를 입력할 시트.


* Naming
    Naming에 필요한 정보를 입력할 시트.


* :doc:`C. Beam Properties <2_c_beam_properties>` / G.Column Properties / Wall Properties
    연결보, 일반기둥, 벽체의 모든 정보를 입력할 시트.


* Output_Naming
    앞에서 입력한 정보들을 바탕으로 이름이 출력되는 시트.


* Output_G.Beam Properties / Output_E.Beam Properties / Output_E.Column Properties
    일반보, 탄성보, 탄성기둥의 모든 정보를 입력할 시트.


* Output_C.Beam Properties / Output_G.Column Properties / Output_Wall Properties
    앞에서 입력한 정보들을 바탕으로 정리된 연결보, 일반기둥, 벽체의 정보가 출력되는 시트.


* Results_C.Beam / Results_G.Beam / Results_E.Beam / Results_G.Column / Results_Wall / Results_E.Column(개별 file)
    해석결과를 바탕으로 연결보, 일반보, 탄성보, 일반기둥, 벽체, 탄성기둥의 강도 검토 결과가 출력되는 시트.

.. raw:: latex

    \newpage
    
.. note::

   Data Conversion Sheets의 셀은 세가지로 분류됩니다.

   .. image:: _static/images/2_DCS_셀_구분.png
      :align: center
      :scale: 80%

   사용자는 하얀색 셀에 모델링 정보를 입력합니다. 
   노란색 셀에도 입력이 가능하지만, PBD_p3d에서 대부분의 내용을 출력해주기 때문에 수정이 필요한 경우에만 입력합니다.

.. toctree::
   :hidden:
   :maxdepth: 1
   :caption: STEPS

   2_naming_rules
   2_etc
   2_materials
   2_nodes_elements
   2_nodal_loads
   2_story_mass
   2_story_data
   2_c_beam_properties
   2_g_column_properties
   2_output_e_g_beam_properties
   2_wall_properties
