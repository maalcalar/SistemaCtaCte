VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMenu 
   BackColor       =   &H8000000D&
   Caption         =   "Sistema de Cuentas Corrientes de Aportaciones"
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   930
   ClientWidth     =   8835
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "frmMenu.frx":030A
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6735
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu m_mae 
      Caption         =   "&Maestro"
      Begin VB.Menu m_mae_usuario 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu m_mae_raya1 
         Caption         =   "-"
      End
      Begin VB.Menu m_mae_gradogrupo 
         Caption         =   "Grados Grupos"
      End
      Begin VB.Menu m_mae_grado 
         Caption         =   "Grados"
      End
      Begin VB.Menu m_mae_regiongrupo 
         Caption         =   "Regiones Grupos"
      End
      Begin VB.Menu m_mae_region 
         Caption         =   "Regiones"
      End
      Begin VB.Menu m_maestro_unidades 
         Caption         =   "Unidades Policiales"
      End
      Begin VB.Menu m_mae_tipsoc 
         Caption         =   "Tipo Socio"
      End
      Begin VB.Menu m_mae_concepto 
         Caption         =   "Concepto de Cobro"
      End
      Begin VB.Menu m_mae_pnp 
         Caption         =   "PNP No Socios"
      End
      Begin VB.Menu m_mae_soc 
         Caption         =   "Socios"
      End
      Begin VB.Menu m_mae_parametro 
         Caption         =   "Parámetros"
      End
      Begin VB.Menu m_mae_directivos 
         Caption         =   "Directivos"
      End
      Begin VB.Menu m_mae_promoc 
         Caption         =   "Promociones"
      End
      Begin VB.Menu m_maestro_x1 
         Caption         =   "-"
      End
      Begin VB.Menu m_salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu m_ele 
      Caption         =   "Elecciones"
      Begin VB.Menu m_ele_padrongral 
         Caption         =   "Padron General"
      End
      Begin VB.Menu m_ele_activos 
         Caption         =   "Padrón Activos Habiles"
      End
      Begin VB.Menu m_ele_unipag 
         Caption         =   "Padrón Por Unidad de Pago"
      End
      Begin VB.Menu m_ele_resumen 
         Caption         =   "Cuadro Resumen Por Tipo"
      End
      Begin VB.Menu m_ele_resumen2 
         Caption         =   "Cuadro Resumen x Tipo Socio"
      End
      Begin VB.Menu m_ele_resxubigeo 
         Caption         =   "Cuadro x UBIGEO"
      End
      Begin VB.Menu m_ele_x1 
         Caption         =   "-"
      End
      Begin VB.Menu m_ele_consulta 
         Caption         =   "Consulta Por Elecciones"
      End
   End
   Begin VB.Menu m_ges 
      Caption         =   "Gestión de Socio"
      Begin VB.Menu m_ges_renuncia 
         Caption         =   "Renuncias"
      End
      Begin VB.Menu m_ges_reingreso 
         Caption         =   "Reingresos"
      End
      Begin VB.Menu m_ges_fracciona 
         Caption         =   "Fraccionamiento"
      End
      Begin VB.Menu m_ges_condona 
         Caption         =   "Condonaciones"
      End
      Begin VB.Menu m_ges_renovac 
         Caption         =   "Renovaciones"
      End
      Begin VB.Menu m_ges_soldscto 
         Caption         =   "Solicitud de Descuentos"
      End
      Begin VB.Menu m_ges_x1 
         Caption         =   "-"
      End
      Begin VB.Menu m_dieco_asig 
         Caption         =   "Asignar Familiares"
      End
      Begin VB.Menu m_ges_glosa 
         Caption         =   "Modificar Glosas TESORERIA"
      End
   End
   Begin VB.Menu m_aporte 
      Caption         =   "Aportes"
      Begin VB.Menu m_apo_estadopagos 
         Caption         =   "Estado de Pagos "
      End
      Begin VB.Menu m_apo_masivo 
         Caption         =   "Estado de Pago Masivo"
      End
      Begin VB.Menu m_apo_estcta 
         Caption         =   "Estado de Cuenta"
      End
   End
   Begin VB.Menu m_dieco 
      Caption         =   "DIECO"
      Begin VB.Menu m_dieco_envio 
         Caption         =   "Enviar Archivo"
      End
      Begin VB.Menu m_dieco_recibir 
         Caption         =   "Recibir Archivo"
      End
      Begin VB.Menu m_dieco_x1 
         Caption         =   "-"
      End
      Begin VB.Menu m_dieco_conxmes 
         Caption         =   "Consulta Socios DIECO x Mes"
      End
      Begin VB.Menu m_dieco_sindscto 
         Caption         =   "Consulta Socios DIECO Sin Dscto"
      End
      Begin VB.Menu m_dieco_conxsoc 
         Caption         =   "Consulta Socios DIECO x Socio"
      End
      Begin VB.Menu m_dieco_socios 
         Caption         =   "Consulta Socios Con Cobro DIECO"
      End
   End
   Begin VB.Menu m_cajamp 
      Caption         =   "Caja-Militar-Policial"
      Begin VB.Menu m_cmp_crear 
         Caption         =   "Enviar Archivo"
      End
      Begin VB.Menu c_cmp_recibir 
         Caption         =   "Recibir Archivo"
      End
      Begin VB.Menu m_cajamp_x1 
         Caption         =   "-"
      End
      Begin VB.Menu m_cmp_conxmes 
         Caption         =   "Consulta Socios CAJMP x Mes"
      End
      Begin VB.Menu m_cmp_socsindscto 
         Caption         =   "Consulta Socios CAJMP Sin Descto"
      End
      Begin VB.Menu m_cmp_conxsoc 
         Caption         =   "Consulta Socios CAJMP x Socio"
      End
      Begin VB.Menu m_cmp_relacion 
         Caption         =   "Consulta Socios conCobro CAJMP"
      End
   End
   Begin VB.Menu m_bcp 
      Caption         =   "BCP"
      Begin VB.Menu m_bcp_envio 
         Caption         =   "Envio BCP"
      End
      Begin VB.Menu m_bcp_retorno 
         Caption         =   "Retorno BCP"
      End
   End
   Begin VB.Menu m_teso 
      Caption         =   "Tesoreria"
      Begin VB.Menu m_teso_cobro 
         Caption         =   "Cobros x Tesoreria"
      End
      Begin VB.Menu m_teso_devol 
         Caption         =   "Devolución de Aportes"
      End
      Begin VB.Menu m_teso_conxfec 
         Caption         =   "Consulta Cobros x Fechas"
      End
      Begin VB.Menu m_teso_conxfecusu 
         Caption         =   "Consulta Cobros x Fechas y Usuario"
      End
   End
   Begin VB.Menu m_repteso 
      Caption         =   "Reportes Tesoreria"
      Begin VB.Menu m_repteso_conxfec 
         Caption         =   "Consulta Cobros x Fechas"
      End
      Begin VB.Menu m_repteso_conxfecusu 
         Caption         =   "Consulta Cobros x Fecha y Usuario"
      End
      Begin VB.Menu m_repteso_cobxsoc 
         Caption         =   "Consulta Cobros x Socio"
      End
      Begin VB.Menu m_repteso_cobxfor 
         Caption         =   "Consulta Cobros x Forma de Pago"
      End
      Begin VB.Menu m_repteso_cobxrec 
         Caption         =   "Consulta Cobros Aportes x Recibo"
      End
      Begin VB.Menu m_repteso_cobxmrecibos 
         Caption         =   "Consulta Cobros Aportes x MRECIBOS"
      End
      Begin VB.Menu m_repteso_cobbancosxrec 
         Caption         =   "Conxulta Cobros Aportes x Bancos"
      End
      Begin VB.Menu m_repteso_cobcarnetxrec 
         Caption         =   "Consulta Cobros Carnet x Recibo"
      End
      Begin VB.Menu m_consulta_asignaciones 
         Caption         =   "Consulta Socios y Asignaciones"
      End
   End
   Begin VB.Menu m_con 
      Caption         =   "Consultas"
      Begin VB.Menu m_con_conxsoc 
         Caption         =   "Saldos x Cobrar x Socio"
      End
      Begin VB.Menu m_con_sdoxcon 
         Caption         =   "Saldos x Cobrar x Concepto"
      End
      Begin VB.Menu m_con_sdoxmes 
         Caption         =   "Saldos x Cobrar x Mes"
      End
      Begin VB.Menu m_con_cobxmes 
         Caption         =   "Cobros x Mes"
      End
      Begin VB.Menu m_consulta_x1 
         Caption         =   "-"
      End
      Begin VB.Menu m_consulta_listadoceo 
         Caption         =   "Listado Alfabetico Para CEOS"
      End
      Begin VB.Menu m_con_socmorosos 
         Caption         =   "Reporte de Socios Activos Morosos"
      End
      Begin VB.Menu m_con_nomorosos 
         Caption         =   "Reporte de Socios Activos Habiles"
      End
      Begin VB.Menu m_con_socios 
         Caption         =   "Reporte de Socios Activos"
      End
      Begin VB.Menu m_con_cuadroresumen 
         Caption         =   "Cuadro Resumen de Aportes"
      End
      Begin VB.Menu m_con_cobxano 
         Caption         =   "Cobranzas x Año"
      End
      Begin VB.Menu m_con_vip 
         Caption         =   "Relación de Socios VIP"
      End
      Begin VB.Menu m_con_ctrldeuda 
         Caption         =   "Control de Deudas x Asociado"
      End
      Begin VB.Menu m_con_socsinapo 
         Caption         =   "Relación Socios Sin Aportes"
      End
      Begin VB.Menu m_con_ressocxtipo 
         Caption         =   "Resumen de Socios x Tipo"
      End
      Begin VB.Menu m_con_socasig 
         Caption         =   "Socios x Asignaciones"
      End
      Begin VB.Menu m_con_dieco 
         Caption         =   "DIECO"
         Begin VB.Menu m_con_dieco_conxmes 
            Caption         =   "Consulta Socios DIECO x Mes"
         End
         Begin VB.Menu m_con_dieco_consindscto 
            Caption         =   "Consulta Socios DIECO Sin Descto"
         End
         Begin VB.Menu m_con_dieco_conxsoc 
            Caption         =   "Consulta Socios DIECO x Socio"
         End
      End
      Begin VB.Menu m_con_cajmp 
         Caption         =   "CAJA MILITAR POLICIAL"
         Begin VB.Menu m_con_cajmp_conxmes 
            Caption         =   "Consulta Socios CAJMP x Mes"
         End
         Begin VB.Menu m_con_cajmp_consindscto 
            Caption         =   "Consulta Socios CAJMP Sin Dscto"
         End
         Begin VB.Menu m_con_cajmp_conxsoc 
            Caption         =   "Consulta Socios CAJMP x Socio"
         End
      End
   End
   Begin VB.Menu m_cont 
      Caption         =   "Contable"
      Begin VB.Menu m_cont_cuadroaportes 
         Caption         =   "Cuadro Anual de Aportes"
      End
      Begin VB.Menu m_cont_cuadrorenova 
         Caption         =   "Cuadro Anual de Renovaciones"
      End
      Begin VB.Menu m_cont_fallecidos 
         Caption         =   "Socios Fallecidos x Año"
      End
      Begin VB.Menu m_cont_renun 
         Caption         =   "Socios Renunciantes x Año"
      End
      Begin VB.Menu m_cont_ing 
         Caption         =   "Socios Ingresantes x Año"
      End
      Begin VB.Menu m_cont_comparaing 
         Caption         =   "Comparativo Ingresantes x Año"
      End
      Begin VB.Menu m_cont_rein 
         Caption         =   "Socios Reingresantes x Año"
      End
   End
   Begin VB.Menu m_servicio 
      Caption         =   "&Servicios"
      Begin VB.Menu m_aporte_crea 
         Caption         =   "Crear Aporte Mensual"
      End
      Begin VB.Menu m_serv_adel 
         Caption         =   "Asignar Adelantos"
      End
      Begin VB.Menu m_serv_actualiza 
         Caption         =   "Actualizar Saldos x Socio"
      End
      Begin VB.Menu m_serv_x1 
         Caption         =   "-"
      End
      Begin VB.Menu m_serv_sdoini 
         Caption         =   "Modificar Saldo Inicial Oct 2017"
      End
      Begin VB.Menu m_serv_glosa 
         Caption         =   "Modificar Glosas TESORERIA"
      End
      Begin VB.Menu m_serv_modifica_codofin 
         Caption         =   "Modificar CODOFIN de Asociado"
      End
      Begin VB.Menu m_serv_estadocta 
         Caption         =   "Estado de Cuenta OCT 2018"
      End
      Begin VB.Menu m_serv_x2 
         Caption         =   "-"
      End
      Begin VB.Menu m_serv_sdorenov 
         Caption         =   "Crear Saldo x Renovaciones"
      End
      Begin VB.Menu m_serv_fotos 
         Caption         =   "Fotos Pruebas"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub m_apo_estado_Click()
   frmAporteEstado.Show
End Sub

Private Sub c_cmp_recibir_Click()
   frmMPRecibe.Show
End Sub

Private Sub m_apo_estadopagos_Click()
   frmAporteEstado.Show
End Sub

Private Sub m_apo_estcta_Click()
   frmAporteEstadoCta.Show
End Sub

Private Sub m_apo_masivo_Click()
   frmAporteMasivo.Show
End Sub


Private Sub m_aporte_crea_Click()
   frmAporteCrea.Show
End Sub

Private Sub m_bcp_envio_Click()
   frmBCPEnvio.Show
End Sub

Private Sub m_bcp_retorno_Click()
   frmBCPRetorno.Show
End Sub

Private Sub m_cmp_conxmes_Click()
   frmMPConxMes.Show
End Sub

Private Sub m_cmp_conxsoc_Click()
   frmMPConxSoc.Show
End Sub

Private Sub m_cmp_crear_Click()
   frmMPEnvio.Show
End Sub

Private Sub m_cmp_relacion_Click()
   frmMPRelacion.Show
End Sub

Private Sub m_cmp_socsindscto_Click()
   frmMPConSinDscto.Show
End Sub

Private Sub m_con_cajmp_consindscto_Click()
   frmMPConSinDscto.Show
End Sub

Private Sub m_con_cajmp_conxmes_Click()
   frmMPConxMes.Show
End Sub

Private Sub m_con_cajmp_conxsoc_Click()
   frmMPConxSoc.Show
End Sub

Private Sub m_con_cobxano_Click()
   frmResumenxAno.Show
End Sub

Private Sub m_con_cobxmes_Click()
   frmConCobxMes.Show
End Sub

Private Sub m_con_conxsoc_Click()
   frmConSdosxSocio.Show
End Sub

Private Sub m_con_ctrldeuda_Click()
   frmConControlDeuda.Show
End Sub

Private Sub m_con_cuadroresumen_Click()
   frmCuadroResumen.Show
End Sub

Private Sub m_con_dieco_consindscto_Click()
   frmDIECOConSinDscto.Show
End Sub

Private Sub m_con_dieco_conxmes_Click()
   frmDIECOConxMes.Show
End Sub

Private Sub m_con_dieco_conxsoc_Click()
   frmDIECOConxSoc.Show
End Sub

Private Sub m_con_nomorosos_Click()
   frmConSocNoMorosos.Show
End Sub

Private Sub m_con_ressocxtipo_Click()
   frmconResSocioxTipo.Show
End Sub

Private Sub m_con_sdoxcon_Click()
   frmConSdosxConcepto.Show
End Sub

Private Sub m_con_sdoxmes_Click()
   frmConSdosxMes.Show
End Sub

Private Sub m_con_socasig_Click()
   frmConAsignaciones.Show
End Sub

Private Sub m_con_socios_Click()
   frmConSocios.Show
End Sub

Private Sub m_con_socmorosos_Click()
   frmConSocMorosos.Show
End Sub

Private Sub m_con_socsinapo_Click()
   frmConSocSinApo.Show
End Sub

Private Sub m_con_vip_Click()
   frmConSocioVIP.Show
End Sub

Private Sub m_consulta_listadoceo_Click()
   frmConListadoCeo.Show
End Sub

Private Sub m_cont_comparaing_Click()
   frmConIngresoxAnoCompara.Show
End Sub

Private Sub m_cont_cuadroaportes_Click()
   frmConAporteAnual.Show
End Sub

Private Sub m_cont_cuadrorenova_Click()
   frmConRenovaAnual.Show
End Sub

Private Sub m_cont_fallecidos_Click()
   frmConFallecidosxAno.Show
End Sub

Private Sub m_cont_ing_Click()
   frmConIngresoxAno.Show
End Sub

Private Sub m_cont_renun_Click()
   frmConRenunciantesxAno.Show
End Sub

Private Sub m_dieco_asig_Click()
   frmMaeAsignado.Show
End Sub

Private Sub m_dieco_conxmes_Click()
   frmDIECOConxMes.Show
End Sub

Private Sub m_dieco_conxsoc_Click()
   frmDIECOConxSoc.Show
End Sub

Private Sub m_dieco_envio_Click()
   frmDiecoEnvio.Show
End Sub

Private Sub m_dieco_recibir_Click()
   frmDiecoRecibe.Show
End Sub

Private Sub m_dieco_sindscto_Click()
   frmDIECOConSinDscto.Show
End Sub

Private Sub m_dieco_socios_Click()
   frmDIECORelacion.Show
End Sub

Private Sub m_ele_activos_Click()
   frmElePadronHabiles.Show
End Sub

Private Sub m_ele_consulta_Click()
   frmEleConsulta.Show
End Sub

Private Sub m_ele_padrongral_Click()
   frmElePadronGral.Show
End Sub

Private Sub m_ele_resumen_Click()
   frmEleResumen.Show
End Sub

Private Sub m_ele_resumen2_Click()
   frmEleResumen2.Show
End Sub

Private Sub m_ele_resxubigeo_Click()
   frmEleResxUbigeo.Show
End Sub

Private Sub m_ele_unipag_Click()
   frmElePadronUnipag.Show
End Sub

Private Sub m_ges_fracciona_Click()
   frmFracMante.Show
End Sub

Private Sub m_ges_glosa_Click()
   frmServGlosas.Show
End Sub

Private Sub m_ges_renuncia_Click()
   frmGesRenuncia.Show
End Sub

Private Sub m_ges_soldscto_Click()
   frmGesSolDscto.Show
End Sub

Private Sub m_mae_directivos_Click()
   frmMaeDirectivo.Show
End Sub

Private Sub m_mae_grado_Click()
   frmMaeGrado.Show
End Sub

Private Sub m_mae_gradogrupo_Click()
   frmMaeGradoGrupo.Show
End Sub

Private Sub m_mae_pnp_Click()
   frmMaePNP.Show
End Sub

Private Sub m_mae_promoc_Click()
   frmMaePromocion.Show
End Sub

Private Sub m_mae_region_Click()
   frmMaeRegion.Show
End Sub

Private Sub m_mae_regiongrupo_Click()
   frmMaeRegionGrupo.Show
End Sub

Private Sub m_mae_soc_Click()
   frmMaeSocio.Show
End Sub

Private Sub m_mae_tipsoc_Click()
   frmMaeTipoSocio.Show
End Sub

Private Sub m_mae_usuario_Click()
   frmMaeUsuario.Show
End Sub

Private Sub m_maestro_unidades_Click()
   frmMaeUnidad.Show
End Sub

Private Sub m_repteso_cobbancosxrec_Click()
   frmCobBancoxRec.Show
End Sub

Private Sub m_repteso_cobcarnetxrec_Click()
   frmCobCarnetxRec.Show
End Sub

Private Sub m_repteso_cobxfor_Click()
   frmCobxFor.Show
End Sub

Private Sub m_repteso_cobxmrecibos_Click()
   frmCobApoxRecMRECIBOS.Show
End Sub

Private Sub m_repteso_cobxrec_Click()
   frmCobApoxRec.Show
End Sub

Private Sub m_repteso_cobxsoc_Click()
   frmCobxSoc.Show
End Sub

Private Sub m_repteso_conxfec_Click()
   frmCobxFec.Show
End Sub

Private Sub m_repteso_conxfecusu_Click()
   frmCobxUsu.Show
End Sub

Private Sub m_salir_Click()
    Unload Me
End Sub

Private Sub m_serv_actualiza_Click()
   frmServActualiza.Show
End Sub

Private Sub m_serv_adel_Click()
   frmServAdel.Show
End Sub

Private Sub m_serv_estadocta_Click()
   frmServEstadoCta.Show
End Sub

Private Sub m_serv_fotos_Click()
   frmMaeFoto.Show
End Sub

Private Sub m_serv_glosa_Click()
   frmServGlosas.Show
End Sub

Private Sub m_serv_modifica_codofin_Click()
   frmServModCodofin.Show
End Sub

Private Sub m_serv_sdoini_Click()
   frmServSdoIni.Show
End Sub

Private Sub m_serv_sdorenov_Click()
   frmServRenovac.Show
End Sub

Private Sub m_teso_bcp_Click()
   frmBCPEnvio.Show
End Sub

Private Sub m_teso_cobbancosxrec_Click()
   frmCobBancoxRec.Show
End Sub

Private Sub m_teso_cobcarnetxrec_Click()
   frmCobCarnetxRec.Show
End Sub

Private Sub m_teso_cobro_Click()
   frmCajaCobros.Show
End Sub

Private Sub m_teso_cobxfor_Click()
   frmCobxFor.Show
End Sub

Private Sub m_teso_cobxmrecibos_Click()
   frmCobApoxRecMRECIBOS.Show
End Sub

Private Sub m_teso_cobxrec_Click()
   frmCobApoxRec.Show
End Sub

Private Sub m_teso_cobxsoc_Click()
   frmCobxSoc.Show
End Sub

Private Sub m_teso_conxfec_Click()
   frmCobxFec.Show
End Sub

Private Sub m_teso_conxfecusu_Click()
   frmCobxUsu.Show
End Sub

Private Sub m_teso_devol_Click()
   frmCajaDevol.Show
End Sub

Private Sub MDIForm_Activate()
   StatusBar1.Panels(1).Text = walmcia + " " + wnomalm
   StatusBar1.Panels(3).Text = "USUARIO " + wcodusu + " " + Trim(wnomusu)
   StatusBar1.Panels(4).Text = wcodcia + " " + Trim(wnomcia)
   StatusBar1.Panels(5).Text = "EJERCICIO " + wanocia

   If Not SUPERVISOR Then
      m_mae_usuario.Enabled = False
   Else
      m_mae_usuario.Enabled = True
   End If
   If MENUMAE = False Then m_mae.Enabled = False Else m_mae.Enabled = True
   If MENUELE = False Then m_ele.Enabled = False Else m_ele.Enabled = True
   If MENUGES = False Then m_ges.Enabled = False Else m_ges.Enabled = True
   If MENUAPO = False Then m_aporte.Enabled = False Else m_aporte.Enabled = True
   If MENUDIE = False Then m_dieco.Enabled = False Else m_dieco.Enabled = True
   If MENUCAJ = False Then m_cajamp.Enabled = False Else m_cajamp.Enabled = True
   If MENUTES = False Then m_teso.Enabled = False Else m_teso.Enabled = True
   If MENUCON = False Then m_con.Enabled = False Else m_con.Enabled = True
   If MENUSER = False Then m_servicio.Enabled = False Else m_servicio.Enabled = True
   If MENUBCP = False Then m_bcp.Enabled = False Else m_bcp.Enabled = True
   If MENURPT = False Then m_repteso.Enabled = False Else m_repteso.Enabled = True
   If MENUCNT = False Then m_cont.Enabled = False Else m_cont.Enabled = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("Esta Seguro de Salir del Sistema ??? :", _
       vbExclamation + vbYesNo + vbDefaultButton2, _
       "Sistema de Cuentas Corrientes de Socios") = vbYes Then
       Cancel = 0
       End
    Else
       Cancel = 1
    End If
End Sub

