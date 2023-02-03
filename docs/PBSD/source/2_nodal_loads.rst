=================
Nodal Loads
=================

Perform-3D에서는 하중을 점(Nodal Loads), 선(Element Loads)의 형식으로만 입력할 수 있습니다.
따라서 Midas Gen의 Floor Loads와 같이 면의 형태로 작용하는 하중은 Nodal Loads, Element Loads로 바꾸어 생성 또는 Import해야합니다.

.. deprecated:: Beta
   기존에는 SDS를 사용하여 Floor Loads를 Nodal Loads로 변환한 후, Nodal Loads를 Import하는 방식을 사용하였습니다.

이러한 Floor Loads 입력의 번거로움을 해결하기 위해, 본 성능기반 내진설계 매뉴얼에서는 하중 대신 반력을 입력하는 방식을 사용합니다.



하중을 Import하는 대신, 모든 수직 부재의 Reaction Force에 음수를 취해서 Import합니다.

필요성
^^^^^^^^^