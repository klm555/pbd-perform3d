==================================
Release Notes
==================================

.. role:: teal
.. role:: under
.. role:: tealbold
.. role:: maroonbold

**2024.05.28**
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
*Version 2.7.0*
~~~~~~~~~~~~~~~~~~~~~

  .. only:: html

    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

  * :maroonbold:`Seismic Design Shear Wall Sheets의 Results_S.Wall_Shear 시트에 벽체 좌표가 추가`\되도록 수정.
  
  .. only:: html

    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

  * :maroonbold:`출력되는 한글(hwp) 파일의 template` 변경.

  .. only:: html

    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

  * 한글(hwp) 파일 출력 시, 한글 파일의 :maroonbold:`그림 속성을 번호 종류 없음으로 변경.`

  .. only:: html

    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

  * 한글(hwp)이나 Word(docx) 파일로 출력 시, :maroonbold:`Preview가 먼저 실행되지 않은 경우 예외를 발생`\시키도록 수정.

  .. only:: html

    .. button-ref:: FIXED
      :ref-type: ref
      :color: info
      :shadow:

  * :maroonbold:`Seismic Design Coupling Beam Sheets의 Design_C.Beam 시트에 데이터가 잘못 입력되는 오류` 수정.

**2024.05.21**
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
*Version 2.6.0*
~~~~~~~~~~~~~~~~~~~~~

  .. only:: html

    .. button-ref:: FEATURE
      :ref-type: ref
      :color: success
      :shadow:

  * 후처리 그래프를 :maroonbold:`한글(hwp) 파일로 출력`\하는 기능 추가.

  .. only:: html

    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

  * 각각의 지진파에 대해, :maroonbold:`모든 위치(ex. 1,5,7,11)의 층간변위비 게이지의 최대값`\을 구한 후, 모든 지진파들의 평균값을 계산하여 층간변위비로 산출하도록 수정.

  .. only:: html

    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

  * :maroonbold:`그래프 사이즈, 제목 등의 디자인 변경`.

  .. only:: html

    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

  * 그래프 미리보기(Preview) 시, :maroonbold:`그래프 제목이 출력되지 않도록` 수정.

  .. only:: html

    .. button-ref:: FIXED
      :ref-type: ref
      :color: info
      :shadow:

  * 벽체의 :maroonbold:`압축변형률, 인장변형률 그래프 크기가 다르게 출력되는 오류`` 수정.

  .. only:: html

    .. button-ref:: FIXED
      :ref-type: ref
      :color: info
      :shadow:

  * Seismic Design Sheets의 입력 여부에 관계없이, :maroonbold:`밑면전단력, 층전단력, 층간변위비의 그래프 미리보기(Preview) 기능이 항상 작동`\되도록 수정.

**2024.02.20**
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
*Version 2.5.3*
~~~~~~~~~~~~~~~~~~~~~

  .. only:: html

    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

  * 연결보 :maroonbold:`Seismic Design Coupling Beam Sheets의 Design_C.Beam 시트`\에 :maroonbold:`Boundary 열이 자동으로 입력`\되도록 수정.

  .. only:: html

    .. button-ref:: FIXED
      :ref-type: ref
      :color: info
      :shadow:

  * 층간변위비 그래프 미리보기(Preview)에서, :maroonbold:`최대고려지진(MCE) 범례에 CP가 아닌 LS로 표기되는 오류` 수정.

**2024.02.15**
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
*Version 2.5.2*
~~~~~~~~~~~~~~~~~~~~~

  .. only:: html

    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

  * 그래프 출력 시, :maroonbold:`그래프 제목이 출력되지 않도록` 수정.

  .. only:: html

    .. button-ref:: FIXED
      :ref-type: ref
      :color: info
      :shadow:

  * 층간변위비 그래프 출력 시, :maroonbold:`최대고려지진(MCE) 범례에 CP가 아닌 LS로 표기되는 오류` 수정.

**2024.02.06**
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
*Version 2.5.1*
~~~~~~~~~~~~~~~~~~~~~

  .. only:: html

    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

  * 연결보 :maroonbold:`Drift Gage 위치(location) 개수 제한 없도록` 수정.

  .. only:: html

    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

  * 연결보 :maroonbold:`Seismic Design Sheets의 수정`\에 따라, :maroonbold:`중력하중 결과도 자동으로 입력`\되도록 수정.

**2024.01.15**
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
*Version 2.4.2*
~~~~~~~~~~~~~~~~~~~~~

  .. only:: html

    .. button-ref:: FEATURE
      :ref-type: ref
      :color: success
      :shadow:

  * 연결보 :maroonbold:`양단고정/1단고정` 설정 기능 추가.

  .. only:: html
   
    .. button-ref:: FIXED
      :ref-type: ref
      :color: info
      :shadow:

  * :maroonbold:`Nodal Loads 값이 모두 Import되지 않는 오류` 수정.

**2023.09.27**
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
*Version 2.4.1*
~~~~~~~~~~~~~~~~~~~~~

 .. only:: html

    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

 * 후처리 시, :maroonbold:`Seismic Design Sheets의 입력`\과 그에 따른 :maroonbold:`그래프 출력`\을 구분하여 작동되도록 수정.

**2023.09.22**
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
*Version 2.4.0*
~~~~~~~~~~~~~~~~~~~~~

 .. only:: html
   
    .. button-ref:: FEATURE
      :ref-type: ref
      :color: success
      :shadow:

 * Seismic Design Sheets에 :maroonbold:`기존값이 있을 시, 자동으로 삭제 후 입력.`

 * :maroonbold:`연결보 전단강도 그래프` 확인 및 출력 가능.

 * Seismic Design Sheets에서 :maroonbold:`해석결과가 없는 부재는 Table 및 Plot 시트에서 자동으로 제거.`

 .. only:: html
   
    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

 * 모든 그래프의 "DE"를 :maroonbold:`"1.2*DBE"`\로 변경

 .. only:: html
   
    .. button-ref:: FIXED
      :ref-type: ref
      :color: info
      :shadow:

 * pdf 연속 출력 시, 엑셀 파일이 닫히지 않고 계속 남아서 출력이 점점 느려지는 사항 수정.

 * 입력값의 앞뒤에 공백이 있는 경우, 자동으로 공백 제거. 

 * 입력창(checkbox, editbox)이 공란인 경우, 예외 처리.

**2023.09.06**
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
*Version 2.3.1*
~~~~~~~~~~~~~~~~~~~~~

 .. only:: html
   
    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

 * 층 분할 시, :maroonbold:`벽체 축변형률(Wall Axial Strain) 게이지가 층 분할을 고려하지 않고 단일 게이지로 Import`\되도록 수정.

 * 그래프 워드(.docx)로 출력할 때 :maroonbold:`해상도(dpi=150)` 개선.
 
 * UI :maroonbold:`글씨체 및 포맷` 변경

**2023.08.31**
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
*Version 2.3.0*
~~~~~~~~~~~~~~~~~~~~~

 .. only:: html
   
    .. button-ref:: FEATURE
      :ref-type: ref
      :color: success
      :shadow:
      
 * :maroonbold:`프로젝트, 건물명 입력창` 추가(pdf 출력용).

 * 메뉴바에 :maroonbold:`계산 시트 경로`\로 바로 이동할 수 있는 메뉴 추가.

 * :maroonbold:`Word(.docx)파일로 출력하는 기능` 추가.

 * Word(.docx)파일 출력 시, overwrite되지 않고 새로운 파일이 생성되도록 함.

 .. only:: html
   
    .. button-ref:: CHANGED
      :ref-type: ref
      :color: warning
      :shadow:

 * 모든 :maroonbold:`DCR 그래프의 x축 limit`\가 3에서 2로 변경됨.

 .. only:: html
   
    .. button-ref:: FIXED
      :ref-type: ref
      :color: info
      :shadow:

 * Nodal Load를 Import할 시, Nodal Load 시트에 비어있는 행이 있는 경우, Error가 발생하는 버그 수정.

 * Analysis Results 파일 선택 시, 아무것도 선택하지 않고 확인을 누르면 기존 경로가 삭제되는 버그 수정.

**2023.08.29**
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
*Version 2.2.0*
~~~~~~~~~~~~~~~~~~~~~

 .. only:: html
   
    .. button-ref:: FEATURE
      :ref-type: ref
      :color: success
      :shadow:

 * :maroonbold:`Release Note, 소프트웨어 정보 확인 창` 추가.

 * Seismic Design Sheets(부재별 결과확인 시트)에서 :maroonbold:`벽체 축변형률(Wall Axial Strain) 결과 확인` 가능.

 * :maroonbold:`연결보 회전각(Beam Rotation) 결과에 Scale Factor 적용` 기능 추가.

 * Seismic Design Sheets의 :maroonbold:`ETC 시트 자동으로 입력.`