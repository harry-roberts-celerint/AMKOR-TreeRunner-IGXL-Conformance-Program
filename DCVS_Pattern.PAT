��!C        �1  �   ()�]	T04UFS-00ultraflexdev 5C:\Users\ultraflexdev\Desktop\Nov12b\DCVS_Pattern.PAT   �                                                                                 �  x�a�DF�T  �U   T  x�c`@�@�.��;@���C璩�X1&"�`��:B��@� )3�               VDDa             VDDaDCVS            VDD_LV_MIPI            VDD_LV_MIPIDCVS            VDD_VS_3_PRB      >      VDD_VS_3_PRBDCVS            VDD_HV_EFUSE      ]      VDD_HV_EFUSEDCVS            VDD_VS_5_PRBa      |      VDD_VS_5_PRBaDCVS            VDD_VS_6_PRB      �      VDD_VS_6_PRBDCVS            VDD_VS_7_PRB      �      VDD_VS_7_PRBDCVS            VDD_HV_VREF      �      VDD_HV_VREFDCVS            
VDD_HV_ADC      �      
VDD_HV_ADCDCVS            VDD_VS_10_PRBa           VDD_VS_10_PRBaDCVS            VDD_HV_IO_ENETa      6     VDD_HV_IO_ENETaDCVS            VDD_HV_IO_QSPI      U     VDD_HV_IO_QSPIDCVS                  VDDa             VDDaDCVS   @        VDD_LV_MIPI            VDD_LV_MIPIDCVS   @        VDD_VS_3_PRB      >      VDD_VS_3_PRBDCVS   @        VDD_HV_EFUSE      ]      VDD_HV_EFUSEDCVS   @        VDD_VS_5_PRBa      |      VDD_VS_5_PRBaDCVS   @        VDD_VS_6_PRB      �      VDD_VS_6_PRBDCVS   @        VDD_VS_7_PRB      �      VDD_VS_7_PRBDCVS   @        VDD_HV_VREF      �      VDD_HV_VREFDCVS   @        
VDD_HV_ADC      �      
VDD_HV_ADCDCVS   @        VDD_VS_10_PRBa           VDD_VS_10_PRBaDCVS   @        VDD_HV_IO_ENETa      6     VDD_HV_IO_ENETaDCVS   @        VDD_HV_IO_QSPI      U     VDD_HV_IO_QSPIDCVS   @              VDDa             VDDaDCVS   �        VDD_LV_MIPI            VDD_LV_MIPIDCVS   �        VDD_VS_3_PRB      >      VDD_VS_3_PRBDCVS   �        VDD_HV_EFUSE      ]      VDD_HV_EFUSEDCVS   �        VDD_VS_5_PRBa      |      VDD_VS_5_PRBaDCVS   �        VDD_VS_6_PRB      �      VDD_VS_6_PRBDCVS   �        VDD_VS_7_PRB      �      VDD_VS_7_PRBDCVS   �        VDD_HV_VREF      �      VDD_HV_VREFDCVS   �        
VDD_HV_ADC      �      
VDD_HV_ADCDCVS   �        VDD_VS_10_PRBa           VDD_VS_10_PRBaDCVS   �        VDD_HV_IO_ENETa      6     VDD_HV_IO_ENETaDCVS   �        VDD_HV_IO_QSPI      U     VDD_HV_IO_QSPIDCVS   �              VDDa             VDDaDCVS   D        VDD_LV_MIPI            VDD_LV_MIPIDCVS   D        VDD_VS_3_PRB      >      VDD_VS_3_PRBDCVS   D        VDD_HV_EFUSE      ]      VDD_HV_EFUSEDCVS   D        VDD_VS_5_PRBa      |      VDD_VS_5_PRBaDCVS   D        VDD_VS_6_PRB      �      VDD_VS_6_PRBDCVS   D        VDD_VS_7_PRB      �      VDD_VS_7_PRBDCVS   D        VDD_HV_VREF      �      VDD_HV_VREFDCVS   D        
VDD_HV_ADC      �      
VDD_HV_ADCDCVS   D        VDD_VS_10_PRBa           VDD_VS_10_PRBaDCVS   D        VDD_HV_IO_ENETa      6     VDD_HV_IO_ENETaDCVS   D        VDD_HV_IO_QSPI      U     VDD_HV_IO_QSPIDCVS   D                   �      �   3     �                                                                                                                                                                	       settle time repeat.
            ___ GATE ON ___
            settle time repeat.
            Apply power to the caps.
        	     Start charging caps for 0.5mS.
        
    First ramp capture.
            Delay .5mS till 2nd capture.
           ��     		repeat 5000
          	> TSet4  00000000000000000000000000000000000 ; 
     		repeat 5000
          	> TSet4  00000000000000000000000000000000000 ; 
     		repeat 5000
          	> TSet4  00000000000000000000000000000000000 ; 
     		repeat 5000
          	> TSet4  00000000000000000000000000000000000 ; 
     		repeat 5000
          	> TSet4  00000000000000000000000000000000000 ; 
     		repeat 5000
          	> TSet4  00000000000000000000000000000000000 ; 
     		repeat 5000
          	> TSet4  00000000000000000000000000000000000 ; 
     		repeat 5000
          	> TSet4  00000000000000000000000000000000000 ; 
     		repeat 5000
          	> TSet4  00000000000000000000000000000000000 ; 
 Second ramp capture.
            Wait 0.5mS
        � Note: the repeat count must be adjusted according the period of ts4.
 So, if the period of ts4 is 100ns, then a repeat count of 5000 results
 in a settling time of 100ns * 5000 = 0.5ms.
                        p           DCVS   VDDaVDD_LV_MIPIVDD_VS_3_PRBVDD_HV_EFUSEVDD_VS_5_PRBaVDD_VS_6_PRBVDD_VS_7_PRBVDD_HV_VREF
VDD_HV_ADCVDD_VS_10_PRBaVDD_HV_IO_ENETaVDD_HV_IO_QSPI            VDDa_pset_1                    MIPI_pset_1                    VS_3_PRB_pset_1                    EFUSE_pset_1                    VS_5_PRBa_pset_1                    VS_6_PRB_pset_1                    VS_7_PRB_pset_1                    VREF_pset_1                    
ADC_pset_1                    VS_10_PRBa_pset_1                    IO_ENETa_pset_1                    IO_QSPI_pset_1                  DCVS_  {  �   MeasureCapsSVM      LVM    SVM       _     c     g     k     |  �   �     �  �   o                                                                                                                                 AllDig  �  #   	AR_CLKP_0    	AR_CLKN_0    PL_1  �  PL_2  �  PL_3  �  PL_4  �  PL_7  �  PH_13  �  PH_4  �  PD_6  w  PG_10  ~  PD_0  o  PD_3  t  PI_2  �  PD_5  v  PH_5  �  NMI  S  PC_11  l  PC_12  m  PC_15  n  PD_2  s  PD_4  u  PG_11    PB_4  h  PB_5  i  PL_0  �  TCK  �  JCOMP  1  TESTMODE  �  RESET_B  �  POR_B  �  POST_AN  �  TMS  �  
XOSC_EXTAL  �  	XOSC_XTAL  �       ����#       �   (_CPIKGenerator_AUTO_GEN_All_Channel_Pin_  � ADC0_CH0     ADC0_CH1    ADC0_CH2    ADC0_CH3    ADC1_CH0    ADC1_CH1    ADC1_CH2    ADC1_CH3    ADC_CH4    ADC_CH5  	  ADC_CH6  
  ADC_CH7    	AR_CLKN_0    	AR_CLKP_0    AR_TXN_0    AR_TXN_1    AR_TXN_2    AR_TXN_3    AR_TXP_0    AR_TXP_1    AR_TXP_2    AR_TXP_3    
ATX_F_DIFF    ATX_F_VI    
ATX_S_DIFF    ATX_S_VI    	CSI1_CLKN    	CSI1_CLKP    CSI1_LANE0N    CSI1_LANE0P    CSI1_LANE1N    CSI1_LANE1P    CSI1_LANE2N     CSI1_LANE2P  !  CSI1_LANE3N  "  CSI1_LANE3P  #  CSI_CLKN  $  CSI_CLKP  %  
CSI_LANE0N  &  
CSI_LANE0P  '  
CSI_LANE1N  (  
CSI_LANE1P  )  
CSI_LANE2N  *  
CSI_LANE2P  +  
CSI_LANE3N  ,  
CSI_LANE3P  -  DCVI_DGS  .  FCCU_F_0  /  FCCU_F_1  0  JCOMP  1  	K_ADC_SRC  2  K_ATX_F  3  K_ATX_PD  4  K_ATX_S  5  K_LFAST_RXC  6  K_LFAST_TERM  7  K_LFAST_TXC  8  K_MTI_LVDS_1  9  K_MTI_LVDS_2  :  
K_REXT_CSI  ;  K_S1_A0  <  K_S1_A1  =  K_S1_EN  >  K_S2_A0  ?  K_S2_A1  @  K_S2_EN  A  	K_S3_Grp1  B  	K_S3_Grp2  C  	K_SAR_CAL  D  K_TMU  E  K_VI_ADJUST  F  	LFAST_RXN  G  	LFAST_RXP  H  	LFAST_TXN  I  	LFAST_TXP  J  MTI_N0  K  MTI_N1  L  MTI_N2  M  MTI_N3  N  MTI_P0  O  MTI_P1  P  MTI_P2  Q  MTI_P3  R  NMI  S  PA_0  T  PA_1  U  PA_10  V  PA_11  W  PA_12  X  PA_13  Y  PA_14  Z  PA_15  [  PA_2  \  PA_3  ]  PA_4  ^  PA_5  _  PA_6  `  PA_7  a  PA_8  b  PA_9  c  PB_0  d  PB_1  e  PB_2  f  PB_3  g  PB_4  h  PB_5  i  PB_6  j  PC_10  k  PC_11  l  PC_12  m  PC_15  n  PD_0  o  PD_1  p  PD_11  q  PD_14  r  PD_2  s  PD_3  t  PD_4  u  PD_5  v  PD_6  w  PD_8  x  PE_13  y  PE_15  z  PF_0  {  PF_14  |  PF_15  }  PG_10  ~  PG_11    PG_5  �  PG_8  �  PG_9  �  PH_11  �  PH_12  �  PH_13  �  PH_4  �  PH_5  �  PH_6  �  PH_7  �  PH_8  �  PH_9  �  PI_0  �  PI_1  �  PI_2  �  PI_4  �  PI_5  �  PI_6  �  PI_7  �  PI_8  �  PI_9  �  PK_0  �  PK_1  �  PK_2  �  PK_3  �  PK_4  �  PK_5  �  PK_6  �  PK_7  �  PL_0  �  PL_1  �  PL_2  �  PL_3  �  PL_4  �  PL_5  �  PL_6  �  PL_7  �  
PMIC_SENSE  �  POR_B  �  POST_AN  �  RESET_B  �  REXT_CSI  �  SAR_ADC_AWG  �  SAR_CAL  �  SAR_CAL2  �  SAR_CAL2_DIFF  �  SAR_CAL_DIFF  �  TCK  �  TESTMODE  �  TMS  �  	TMU_ANODE  �  TMU_ANODE_DIFF  �  TMU_ANODE_IO  �  TMU_CATHODE  �  TMU_CATHODE_DIFF  �  TMU_CATHODE_IO  �  
VDD_HV_ADC  �  VDD_HV_EFUSE  �  VDD_HV_FIRC  �  VDD_HV_FXOSC  �  
VDD_HV_IO1  �  
VDD_HV_IO2  �  VDD_HV_IO_AUR  �  VDD_HV_IO_ENETa  �  VDD_HV_IO_ENETb  �  VDD_HV_IO_LFAST  �  VDD_HV_IO_MTI  �  VDD_HV_IO_QSPI  �  VDD_HV_LFASTPLL  �  VDD_HV_MIPI  �  
VDD_HV_PLL  �  VDD_HV_PLL_AUR  �  
VDD_HV_TMU  �  VDD_HV_VREF  �  VDD_LV_MIPI  �  
VDD_LV_PLL  �  VDD_LV_PLL_AUR  �  VDD_VS_10_PRBa  �  VDD_VS_10_PRBb  �  VDD_VS_3_PRB  �  VDD_VS_5_PRBa  �  VDD_VS_5_PRBb  �  VDD_VS_6_PRB  �  VDD_VS_7_PRB  �  VDDa  �  VDDb  �  VDDc  �  VDDd  �  
VREFH_ADC0  �  VREFH_ADC0_DIFF  �  
VREFH_ADC1  �  VREFH_ADC1_DIFF  �  VSS  �  
XOSC_EXTAL  �  	XOSC_XTAL  �  AllDig  � #   	AR_CLKP_0	AR_CLKN_0PL_1PL_2PL_3PL_4PL_7PH_13PH_4PD_6PG_10PD_0PD_3PI_2PD_5PH_5NMIPC_11PC_12PC_15PD_2PD_4PG_11PB_4PB_5PL_0TCKJCOMPTESTMODERESET_BPOR_BPOST_ANTMS
XOSC_EXTAL	XOSC_XTAL
DCVS_PINSa  �    VDDaVDD_LV_MIPIVDD_VS_3_PRBVDD_HV_EFUSEVDD_VS_5_PRBaVDD_VS_6_PRBVDD_VS_7_PRBVDD_HV_VREF
VDD_HV_ADCVDD_VS_10_PRBaVDD_HV_IO_ENETaVDD_HV_IO_QSPI
DCVS_PINSb  �    VDDbVDD_VS_5_PRBbVDD_VS_10_PRBbVDD_HV_IO_ENETb
DCVS_PINSc  �    VDDc
DCVS_PINSd  �    VDDdIO_PINS  � #   	AR_CLKP_0	AR_CLKN_0PL_1PL_2PL_3PL_4PL_7PH_13PH_4PD_6PG_10PD_0PD_3PI_2PD_5PH_5NMIPC_11PC_12PC_15PD_2PD_4PG_11PB_4PB_5PL_0TCKJCOMPTESTMODERESET_BPOR_BPOST_ANTMS
XOSC_EXTAL	XOSC_XTALLeakPins  � #   	AR_CLKP_0	AR_CLKN_0PL_1PL_2PL_3PL_4PL_7PH_13PH_4PD_6PG_10PD_0PD_3PI_2PD_5PH_5NMIPC_11PC_12PC_15PD_2PD_4PG_11PB_4PB_5PL_0TCKJCOMPTESTMODERESET_BPOR_BPOST_ANTMS
XOSC_EXTAL	XOSC_XTAL        
       	   �� �  �  �  �                                                                                                                                                                          HSDMQ  single    Pinmap ;C:\Users\ultraflexdev\Desktop\Nov12b\8.10.14_CelerintC.xlsm  5C:\Users\ultraflexdev\Desktop\Nov12b\DCVS_Pattern.PAT
   allow_differential allow_mixed_1x2x  allow_mto_dash allow_scan_in_srm  	cdl_cache check_set_msb check_vm_min_size init_pattern   time_domain   vm_block_size                #              AllDig#       #   	AR_CLKP_0	AR_CLKN_0PL_1PL_2PL_3PL_4PL_7PH_13PH_4PD_6PG_10PD_0PD_3PI_2PD_5PH_5NMIPC_11PC_12PC_15PD_2PD_4PG_11PB_4PB_5PL_0TCKJCOMPTESTMODERESET_BPOR_BPOST_ANTMS
XOSC_EXTAL	XOSC_XTAL                       AllDig#   #                                $   (   ,   0   4   8   <   @   D   H   L   P   T   X   \   `   d   h   l   p   t   x   |   �   �   �   #   	AR_CLKP_0	AR_CLKN_0PL_1PL_2PL_3PL_4PL_7PH_13PH_4PD_6PG_10PD_0PD_3PI_2PD_5PH_5NMIPC_11PC_12PC_15PD_2PD_4PG_11PB_4PB_5PL_0TCKJCOMPTESTMODERESET_BPOR_BPOST_ANTMS
XOSC_EXTAL	XOSC_XTAL0      ����                 TSet4           DCVS     VDDa             VDDa             DCVS     VDD_LV_MIPI            VDD_LV_MIPI            DCVS     VDD_VS_3_PRB      >      VDD_VS_3_PRB   >         DCVS     VDD_HV_EFUSE      ]      VDD_HV_EFUSE   ]         DCVS     VDD_VS_5_PRBa      |      VDD_VS_5_PRBa   |         DCVS     VDD_VS_6_PRB      �      VDD_VS_6_PRB   �         DCVS     VDD_VS_7_PRB      �      VDD_VS_7_PRB   �         DCVS     VDD_HV_VREF      �      VDD_HV_VREF   �         DCVS     
VDD_HV_ADC      �      
VDD_HV_ADC   �         DCVS     VDD_VS_10_PRBa           VDD_VS_10_PRBa           DCVS     VDD_HV_IO_ENETa      6     VDD_HV_IO_ENETa   6        DCVS     VDD_HV_IO_QSPI      U     VDD_HV_IO_QSPI   U           DCVS`.  D     MeasureCaps�  �   �            LVM    SVM   �)  �  �+             �+    �,  �  �1     F.     J.     N.     [)  �   �1      \.                                                                                                                          