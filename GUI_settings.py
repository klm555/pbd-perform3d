from PyQt5.QtCore import QSettings, QCoreApplication

class setting_var():

# setting에 저장된 value를 불러와서 입력
        # QSettings 클래스 생성
        QCoreApplication.setOrganizationName('CNP_Dongyang')
        QCoreApplication.setApplicationName('PBD_with_PERFORM-3D')
        self.setting = QSettings()

        # setting에 저장된 value를 불러와서 입력
        self.setting.beginGroup('file_path')
        self.data_conv_path_editbox.setText(self.setting.value('data_conversion_file_path', 'C:\\'))
        self.data_conv_path_editbox_2.setText(self.setting.value('data_conversion_file_path', 'C:\\'))
        self.data_conv_path_editbox_3.setText(self.setting.value('data_conversion_file_path', 'C:\\'))
        self.result_path_editbox.setText(self.setting.value('result_file_path', 'C:\\'))
        self.setting.endGroup()
        
        # Tab 1
        self.setting.beginGroup('setting_tab1')
        self.import_node_checkbox.setChecked(self.setting.value('import_node', True, type=bool))
        self.import_beam_checkbox.setChecked(self.setting.value('import_beam', True, type=bool))
        self.import_col_checkbox.setChecked(self.setting.value('import_column', False, type=bool))
        self.import_wall_checkbox.setChecked(self.setting.value('import_wall', True, type=bool))
        self.import_plate_checkbox.setChecked(self.setting.value('import_plate', False, type=bool))
        self.import_WR_gage_checkbox.setChecked(self.setting.value('import_wall_rotation_gage', True, type=bool))
        self.import_WAS_gage_checkbox.setChecked(self.setting.value('import_wall_axial_strain_gage', True, type=bool))
        self.import_mass_checkbox.setChecked(self.setting.value('import_mass', True, type=bool))
        self.import_nodal_load_checkbox.setChecked(self.setting.value('import_nodal_load', True, type=bool))
        self.DL_name_editbox.setText(self.setting.value('DL_name', 'DL'))
        self.LL_name_editbox.setText(self.setting.value('LL_name', 'LL'))
        self.drift_pos_editbox.setText(self.setting.value('drift_positions', '2,5,7,11'))
        self.convert_wall_checkbox.setChecked(self.setting.value('convert_wall', True, type=bool))
        self.convert_cbeam_checkbox.setChecked(self.setting.value('convert_beam', True, type=bool))
        self.convert_gcol_checkbox.setChecked(self.setting.value('convert_column', True, type=bool))
        self.setting.endGroup()
        
        # Tab 3
        self.setting.beginGroup('setting_tab3')
        self.base_SF_checkbox.setChecked(self.setting.value('base_SF', True, type=bool))
        self.story_SF_checkbox.setChecked(self.setting.value('story_SF', False, type=bool))
        self.IDR_checkbox.setChecked(self.setting.value('IDR', True, type=bool))
        self.BR_checkbox.setChecked(self.setting.value('BR', True, type=bool))
        self.E_BSF_checkbox.setChecked(self.setting.value('E_BSF', False, type=bool))
        self.E_CSF_checkbox.setChecked(self.setting.value('E_CSF', False, type=bool))
        self.WAS_checkbox.setChecked(self.setting.value('WAS', True, type=bool))
        self.WR_checkbox.setChecked(self.setting.value('WR', True, type=bool))
        self.WSF_checkbox.setChecked(self.setting.value('WSF', True, type=bool))
        self.WSF_each_checkbox.setChecked(self.setting.value('WSF_each', True, type=bool))
        self.bldg_name_editbox.setText(self.setting.value('bldg_name', '101동'))
        self.story_gap_editbox.setText(self.setting.value('story_gap', '2'))
        self.max_shear_editbox.setText(self.setting.value('max_shear', '60000'))
        self.setting.endGroup()