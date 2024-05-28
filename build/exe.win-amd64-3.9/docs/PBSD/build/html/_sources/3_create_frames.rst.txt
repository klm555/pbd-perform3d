
Frame 생성
========================

Frame은 Midas Gen의 Active와 비슷한 역할을 합니다. 노드와 부재가 많은 경우, **Frame을 만들어서 필요한 노드와 부재만 Active하여 모델링이 가능합니다.**
그러나 Midas Gen의 Active 기능에 비해 매우 불편하고 Frame이 없이도 모델링이 불가능하진 않기 때문에, 사용자에 따라 필요한 Frame만 설정하여 사용하는 경우도 있습니다.
또한 건물의 규모가 작다면 Frame을 사용하지 않고 모든 노드와 부재에 접근할 수도 있습니다.
그럼에도 불구하고 Frame을 사용하는 것이 다방면에서 편리하기 때문에, Frame을 사용하는 것을 권장합니다.

본 업무절차서에서는 모든 부재의 Frame을 생성하여 모델링하는 방법을 소개합니다.

Frame 생성 및 이름 입력
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

이름만 입력된 비어있는 Frame을 먼저 생성한 후, 비어있는 Frame에 노드와 부재를 추가합니다.

.. topic:: What to do

   1. :guilabel:`Add or Delete Frames`\를 클릭합니다. 생성된 창에서 :guilabel:`New`\를 클릭합니다.

      .. figure:: _static/images/3_create_frames_name.gif
         :align: center

      Data Conversion Sheets의 Input_Naming 시트를 확인합니다. 이 시트의 ``Frame`` 열에 생성된 가장 첫 열, 첫 행의 이름을 Perform-3D에 입력합니다.

      .. figure:: _static/images/3_create_frames_Naming_시트.png
         :align: center
         :scale: 80%

      입력한 후에 :guilabel:`OK`\를 클릭하면 비어있는 Frame이 생성됩니다.

   2. 같은 방법으로 ``Frame`` 열에 생성된 이름을 첫번쨰 열부터 차례대로 모두 입력합니다. 다만, 모델링에 필요없다고 판단되는 경우에는 생성하지 않아도 됩니다.

   3. 모든 Frame이 생성되면, Frame에 해당하는 노드와 부재를 추가하기 위해 View 창에서 :guilabel:`Plan`\을 클릭합니다.

.. tip::
   Frame의 이름을 입력하는 과정은 어렵지 않으나, 생성해야할 Frame의 수가 많은 경우에는 많은 시간, 노동, 실수가 동반될 수 있습니다. 
   이를 방지하기 위해, Frame의 이름을 입력하는 과정을 자동화하는 방법을 소개합니다. 
   이 방법은 이 후에 진행될 Section 이름 입력, Property Check & Save 과정 등에서도 활용이 가능합니다.

   Frame의 이름을 자동으로 입력하기위해 사용할 소프트웨어는 **key_macro** 입니다. 
   마우스 또는 키보드의 조작 방법과 순서를 명령문으로 기록하면, **인간이 직접 조작하는 것과 동일하게 key_macro가 마우스와 키보드를 조작합니다.**

   1. key_macro를 실행합니다.

   2. 먼저 매크로 속도를 변경합니다. 매크로 속도가 너무 빠르면 오류가 자주 생기고, 너무 느리면 시간이 오래 걸리기 때문에 적절한 속도를 찾아야 합니다.
      매크로 속도는 :guilabel:`설정`\을 클릭하고 매크로 이벤트 실행 주기를 변경하여 설정할 수 있습니다. 70밀리초로 설정하는 것을 권장드리고, 사용자의 필요에 따라 변경합니다.

      70밀리초를 입력한 후에 :guilabel:`OK`\을 클릭합니다.

      .. figure:: _static/images/3_create_frames_key_macro_speed.png
         :align: center
         :scale: 80%

   3. 매크로 명령문 작성을 위해 :guilabel:`추가`\을 클릭합니다. 
      새로 생성된 창에서 매크로 이름을 입력합니다. 프로젝트 또는 건물마다 매크로 명령문을 재작성하기 번거롭기 때문에, 재사용을 위해 사용자가 기억할 수 있도록 이름을 지정하는 것이 좋습니다.

      .. figure:: _static/images/3_create_frames_key_macro_name.png
         :align: center
         :scale: 80%

   4. Data Conversion Sheets의 Input_Naming 시트와 Perform-3D(또는 원격데스크톱의 Perform-3D)를 각각의 모니터에 띄웁니다.    

      .. figure:: _static/images/3_create_frames_dual_monitor.png
         :align: center

         *듀얼 모니터를 사용한다고 가정합니다. 방식은 싱글 모니터에서도 동일합니다.*

   사용자가 직접 엑셀의 이름을 복사하고 Perform-3D에 붙여넣는 일련의 순서를 조금 더 자세히 나열해 보겠습니다.

    (엑셀 창 클릭) :octicon:`arrow-right` 엑셀에서 이름 :kbd:`Ctrl` + :kbd:`C` :octicon:`arrow-right` (Perform-3D 창 클릭) 
    :octicon:`arrow-right` Perform-3D에서 :guilabel:`New` 버튼 클릭 :octicon:`arrow-right` 이름 :kbd:`Ctrl` + :kbd:`V` 
    :octicon:`arrow-right` Perform-3D에서 :guilabel:`OK` 버튼 클릭 
    
    :octicon:`arrow-right` 입력할 Frame 개수만큼 반복
   
   위의 과정을 매크로 명령문에 순서대로 작성합니다.

   5. Perform-3D를 원격으로 실행하는 경우, 엑셀 창 또는 원격데스크톱 창을 클릭하여 활성화하는 과정이 없으면 매크로 명령문이 정상적으로 작동하지 않습니다.
      따라서 엑셀을 먼저 활성화하기 위해 Data Conversion Sheets의 제목표시줄을 클릭하는 명령을 작성합니다.

      :guilabel:`마우스 추가`\를 클릭합니다. 
      마우스를 움직이면 새로 생성된 창에 마우스의 좌표가 실시간으로 표시되고, :kbd:`F10`\을 누르면 마우스의 좌표가 기록됩니다.

      .. figure:: _static/images/3_create_frames_key_macro_click_window_bar.png
         :align: center
         :scale: 40%

      마우스를 제목표시줄에 올리고 :kbd:`F10`\을 눌러 마우스의 좌표를 기록됩니다.

      .. figure:: _static/images/3_create_frames_key_mouse_capture.png
         :align: center
         :scale: 80%

      마우스 버튼에서 왼쪽 버튼, 클릭을 누릅니다. 마우스 왼쪽 버튼으로 해당 좌표를 클릭하는 명령을 추가하기 위함입니다.
      :guilabel:`OK`\를 눌러 명령 추가를 완료합니다.

      .. figure:: _static/images/3_create_frames_key_mouse_capture_succeed.png
         :align: center
         :scale: 80%

      해당 명령이 위와 같이 추가되었음을 알 수 있습니다.

   6. 다음으로 엑셀의 이름값을 :kbd:`Ctrl` + :kbd:`C`\하기 위한 명령을 추가합니다.

      :kbd:`Ctrl` + :kbd:`C`\를 누르는 과정은 아래와 같이 다시 세분화할 수 있습니다.

       * :kbd:`Ctrl` 누르기
       * :kbd:`C` 누르고 떼기
       * :kbd:`Ctrl` 떼기

      이 명령을 추가하기 위해 :guilabel:`키보드 추가`\를 클릭합니다. 

      .. figure:: _static/images/3_create_frames_key_keyboard_capture.png
         :align: center
         :scale: 80%

      먼저 ":kbd:`Ctrl` 누르기"를 추가하기 위해 위와 같이 :kbd:`Ctrl`\과 누르기를 선택한 후 :guilabel:`OK`\를 누릅니다.

      .. figure:: _static/images/3_create_frames_key_keyboard_capture_2.png
         :align: center
         :scale: 80%

      같은 방법으로  ":kbd:`C` 누르고 떼기"를 추가하기 위해 위와 같이 :kbd:`C`\와 누르고 떼기를 선택한 후 :guilabel:`OK`\를 누릅니다.

      마지막으로 ":kbd:`Ctrl` 떼기"를 추가하기 위해 위와 같이 :kbd:`Ctrl`\과 떼기를 선택한 후 :guilabel:`OK`\를 누릅니다.

      .. figure:: _static/images/3_create_frames_key_keyboard_capture_succeed.png
         :align: center
         :scale: 80%

      추가한 3개의 명령이 차례대로 추가되었음을 알 수 있습니다.

   7. Perform-3D 창을 클릭하는 명령을 추가하기 전에, 다음 부재의 입력이 쉽도록 미리 다음 셀로 이동하는 명령을 추가합니다.
      다음 부재는 현재 셀의 아래에 있으므로 :kbd:`↓`\를 눌러야합니다.

      6번과 같은 방법으로 ":kbd:`↓`\ 누르고 떼기"를 추가합니다.

      .. figure:: _static/images/3_create_frames_key_keyboard_capture_3.png
         :align: center
         :scale: 80%

   8. 5,6,7번을 참고하여 모든 명령을 추가합니다. 모든 명령이 추가된 매크로 명령문은 아래와 같습니다.

      .. figure:: _static/images/3_create_frames_key_macro_comment.png
         :align: center
         :scale: 80%

      :guilabel:`문자열 추가`\의 주석 기능을 이용하여 위와 같이 각 명령어에 코멘트를 추가할 수 있습니다. 
      명령문이 길어지면 어떤 명령어가 어떤 기능을 하는지 알기 어려우므로, 코멘트를 추가하는 것을 권장합니다.

   9. **key_macro** 에는 입력된 횟수만큼 명령문을 반복할 수 있는 기능이 있습니다. 
      매크로 반복 실행 횟수에 해당 열의 개수를 입력하면, 해당 열의 모든 부재명을 자동으로 입력하도록 할 수 있습니다.

      .. figure:: _static/images/3_create_frames_key_macro_iter_num.png
         :align: center

   10. key_macro에서 매크로 명령문의 실행과 중지는 단축키로만 가능합니다.
       :guilabel:`시작/중지 조건`\을 클릭하고, 시작 단축키, 중지 단축키를 설정합니다. 
       본 업무절차서에서는 :kbd:`F2`\을 시작 단축키로, :kbd:`Esc`\를 중지 단축키로 설정하였습니다.

       .. figure:: _static/images/3_create_frames_key_macro_start_stop.png
         :align: center
         :scale: 80%
       
       설정 완료 후 :guilabel:`OK`\를 누릅니다.
       다시 한 번 :guilabel:`OK`\를 눌러 매크로 추가를 완료합니다.

   11. 작성한 매크로 명령문을 향후에도 사용할 수 있도록 :guilabel:`저장`\을 누릅니다.

   12. "매크로 실행 가능"을 클릭하기 전까지는 시작 단축키와 중지 단축키를 눌러도 매크로가 실행되지 않습니다.
       매크로를 실행하기 위해 "매크로 실행 가능"을 클릭합니다.

       :kbd:`F2` 또는 사용자가 설정한 시작 단축키를 누르면 아래와 같이 매크로가 실행됩니다.

       .. figure:: _static/images/3_create_frames_macro.gif
          :align: center

Frame에 부재 추가
^^^^^^^^^^^^^^^^^^^^^^^^^^^

앞서 생성한 Frame에 해당되는 이름의 부재를 추가합니다.

.. topic:: What to do

   1. Perform-3D 창에서 부재를 추가하려는 Frame을 선택합니다.

      Add Nodes 탭을 선택합니다. Add Nodes 탭은 해당 프레임에 노드를 추가할 때 사용합니다.
      반대로 Delete Nodes 탭은 해당 프레임에서 노드를 삭제할 때 사용합니다.

      .. figure:: _static/images/3_create_frames_choose_frame.png
         :align: center

   2. View 창의 :guilabel:`Plan`\을 클릭합니다.
   
      해당 건물의 평면도/탄성설계 모델을 확인하여 Perform-3D에서 해당 부재의 위치를 찾습니다. 부재의 위치가 확인되면 해당 부재를 구성하는 노드를 마우스 드래그로 선택합니다.

      .. figure:: _static/images/3_create_frames_name.gif
         :align: center

      Perform-3D에서는 모델의 모든 노드가 보여지기 때문에, 추가하려는 노드의 위치를 찾기 어려울 수 있습니다. 노드의 정확한 위치를 확인하기 위해서는 View를 변경해가며 모델을 

      Data Conversion Sheets의 Input_Naming 시트를 확인합니다. 이 시트의 ``Frame`` 열에 생성된 가장 첫 열, 첫 행의 이름을 Perform-3D에 입력합니다.

Frame에 부재를 추가하는 작업은 앞서 소개된 Frame을 생성하고 이름을 입력하는 작업과는 달리 자동화가 되어있

기타 Frame 추가
^^^^^^^^^^^^^^^^^^^^

각각의 부재의 Frame 외에도 사용자의 편의에 따라 추가적으로 Frame을 생성할 수 있습니다.
주로 많이 사용되는 Frame은 층별, 구간별 Frame이 있습니다.




   

   