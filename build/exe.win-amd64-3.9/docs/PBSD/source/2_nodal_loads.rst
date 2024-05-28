=================
Nodal Loads
=================

Perform-3D에서는 하중을 점(Nodal Loads), 선(Element Loads)의 형식으로만 입력할 수 있습니다.
따라서 Midas Gen의 Floor Loads와 같이 면의 형태로 작용하는 하중은 Nodal Loads, Element Loads로 바꾸어 생성 또는 Import해야합니다.

.. deprecated:: Beta
   기존에는 SDS를 사용하여 Floor Loads를 Nodal Loads로 변환한 후, Nodal Loads를 Import하는 방식을 사용하였습니다.

이러한 Floor Loads 입력의 번거로움을 해결하기 위해, 본 성능기반 내진설계 업무절차서에서는 **하중 대신 반력**\을 입력하는 방식을 사용합니다.
점, 선, 면의 다양한 형태로 입력된 모든 하중에 대한 반력은 수직 부재의 Nodes에서만 발생하기 때문에, 
반력을 이용한다면 모든 하중을 Nodal Loads의 형태로 입력할 수 있습니다.

Supports 생성
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. topic:: What to do

   1. Midas Gen에서 반력은 Supports가 설정된 Nodes에서만 생성됩니다. 
      따라서 모든 수직 부재에서의 반력을 구하기 위해서는, 모든 수직 부재들에 Supports를 생성해야 합니다.
      우선, 모든 수직 부재를 선택하기 위해 :guilabel:`Select Nodes by Identifying...`\를 클릭합니다.
      
      .. figure:: _static/images/2_select_vertical_elem.gif
         :align: center

   2. 생성된 창에서 :guilabel:`Select Type` - :guilabel:`Material`\을 선택한 후, Nodes만 선택하기 위해 :guilabel:`Nodes`\만 체크합니다.

      .. image:: _static/images/2_select_vertical_elem_2.gif
         :align: center

      Wall, Column과 같은 수직 부재의 재료들만 선택한 후, :kbd:`Add` - :kbd:`Close`\를 클릭합니다.

   3. 선택된 수직 부재들에 Supports를 설정해주기 위해 :guilabel:`Boundary` - :guilabel:`Define Supports`\를 클릭합니다.

      .. figure:: _static/images/2_set_vertical_elem_support.gif
         :align: center

      생성된 창에서 :guilabel:`Dz`\만 체크 후 :kbd:`Apply`\를 클릭하면, 선택된 모든 수직 부재들에 Supports가 생성됩니다.

반력 입력
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

   1. 반력을 확인하기 위해서는 반드시 해석을 실시해야 합니다. 해석을 진행합니다.
   
   2. Supports의 반력을 확인하기 위해, :guilabel:`Results` - :guilabel:`Results Table` - :guilabel:`Reaction`\를 클릭합니다.


      :doc:`하중의 의미 <2_etc>`\를 다시 한 번 확인하고, 생성된 창에서 DL과 LL에 포함되는 하중을 모두 체크한 후 :kbd:`OK`\를 클릭합니다.