var ruta;
ruta="imagenes/";

stm_bm(["menu3ba5",810,"",ruta+"blank.gif",0,"","",0,0,200,0,1000,1,0,0,"","100%",0,0,1,2,"default","hand",""],this);
stm_bp("p0",[0,4,0,0,3,5,21,0,100,"",-2,"",-2,50,0,0,"#999999","#E6EFF9",ruta+"bg_02.gif",3,0,0,"#000000","",-1,-1,0,"#FFFFF7","",3,"bg_03.gif",37,3,0,"#FFFFF7","",3,"",-1,-1,0,"#FFFFF7","",3,"bg_01.gif",37,3,0,"#FFFFF7","",3,"","","","",0,0,0,0,0,0,0,0]);
stm_ai("p0i0",[0,"Inicio","","",-1,-1,0,"Principal.asp","_self","","",ruta+"folder_home.png",ruta+"folder_home_ov.png",20,20,0,"","",0,0,0,0,1,"#E6EFF9",1,"#E6EFF9",1,"","",3,3,0,0,"#E6EFF9","#000000","#FFFFFF","#DFA040","bold 8pt Verdana","italic bold 8pt Verdana",0,0],100,30);
stm_aix("p0i1","p0i0",[0,"SICON","","",-1,-1,0,"","_self","","",ruta+"document_graph.png",ruta+"document_graph_ov.png",20,20,0,"","",0,0,0,1,1,"#E6EFF9",1,"#E6EFF9",1,"","4545454.gif",3,3,0,0,"#E6EFF9","#000000","#FFFFFF","#FFFFFF","bold 8pt Verdana","bold 8pt Verdana"],100,30);
stm_bpx("p1","p0",[1,4,0,2,2,5,0,10,100,"",-2,"",-2,50,2,3,"#333333","#333333","",3,1,1,"#000000","",-1,-1,0,"#FFFFF7","",3,"",-1,-1,0,"#FFFFF7","",3,"",-1,-1,0,"#FFFFF7","",3,"",-1,-1]);
stm_aix("p1i0","p0i1",[0,"Directorio","","",-1,-1,0,"Directorio.asp","_self","","","","",0,0,0,"","",0,0,0,0,1,"#E6EFF9",1,"#000000",0,"","",3,3,0,0,"#E6EFF9","#000000","#FFFFFF","#DFA040"],165,20);
stm_aix("p1i1","p1i0",[0,"EE.FF.","","",-1,-1,0,"","_self","","","","",0,0,0,ruta+"arrowgrey-r.gif",ruta+"arrowgrey-r.gif",10,5],165,20);
stm_bp("p2",[1,2,-20,2,0,5,0,0,80,"",-2,"",-2,50,2,3,"#333333","#333333","",3,1,1,"#000000"]);
stm_aix("p2i0","p1i0",[0,"Balance General","","",-1,-1,0,"BalanceGeneral.asp"],220,0);
stm_aix("p2i1","p1i0",[0,"Ganancias y Pérdidas","","",-1,-1,0,"EstGananyPerdidas.asp"],220,0);
stm_aix("p2i2","p1i0",[0,"Flujo Efectivo","","",-1,-1,0,"EstFlujoEfectivo.asp"],220,20);
stm_aix("p2i3","p1i0",[0,"Cambio Patrimonio","","",-1,-1,0,"CambioPatrimonio.asp"],220,0);
stm_aix("p2i4","p1i0",[0,"Dividendos Declarados","","",-1,-1,0,"dividendos.asp"],220,0);
stm_ep();
stm_aix("p1i2", "p1i1", [0, "Consultas Dinámicas"], 165, 0);
stm_bpx("p3", "p2", []);
stm_aix("p3i0", "p1i0", [0, "Dividendos Declarados", "", "", -1, -1, 0, "dividendos_p.asp"], 220, 0);
stm_ep();
stm_aix("p1i3", "p1i1", [0, "Anexos"], 165, 0);
stm_bpx("p3","p2",[]);
stm_aix("p3i0", "p1i0", [0,"Anexo 2 - Existencias / bienes realizables","","",-1,-1,0,"Anexo2.asp"],220,0);
stm_aix("p3i1", "p1i0", [0,"Anexo 5 - Inmuebles, Maquinaria Y Equipo","","",-1,-1,0,"Anexo5.asp"],220,0);
stm_aix("p3i2", "p1i0", [0,"Anexo 6 - Depreciación Acumulada De Inmuebles, Maquinaria Y Equipo","","",-1,-1,0,"Anexo6.asp"],220,0);
stm_aix("p3i3", "p1i0", [0,"Anexo 7 - Activos Intangibles Y Otros Activos","","",-1,-1,0,"Anexo7.asp"],220,0);
stm_ep();
stm_aix("p1i3","p1i1",[0,"Sistema Intermedio","","",-1,-1,0,"SistemaIntermedio.asp"],165,20);
stm_bpx("p4","p2",[]);
stm_aix("p4i0","p1i0",[0,"Balance de Comprobación","","",-1,-1,0,"BalanceComprobacion.asp"],220,0);
stm_aix("p4i1","p1i0",[0,"CSI","","",-1,-1,0,"BalanceCSI.asp"],220,0);
stm_ep();
stm_aix("p1i3","p1i1",[0,"Consistencias","","",-1,-1,0,"SistemaIntermedio.asp"],165,20);
stm_bpx("p4","p2",[]);
stm_aix("p4i0","p1i0",[0,"Estado de situación financiera","","",-1,-1,0,"Consistencia_esf.asp"],220,0);
stm_aix("p4i1","p1i0",[0,"Estado de resultados","","",-1,-1,0,"Consistencia_er.asp"],220,0);
stm_ep();
stm_ep();
stm_aix("p0i2","p0i0",[0,"FONAFE","","",-1,-1,0,"","_self","","",ruta+"gobierno2.jpg",ruta+"gobierno2.jpg",21],100,30);
stm_bpx("p5","p2",[1,4,0,2,2,5,0,0,100]);
stm_aix("p5i0","p1i0",[0,"Directorio","","",-1,-1,0,"Directorio_f.asp"],165,0);
stm_aix("p5i1","p1i0",[0,"Balance General","","",-1,-1,0,"BGFonafe.asp"],165,0);
stm_aix("p5i2","p1i0",[0,"Presupuesto","","",-1,-1,0,"Presupuesto.asp"],165,0);
stm_aix("p5i3","p1i0",[0,"Flujo de Caja","","",-1,-1,0,"FlujoCaja.asp"],165,0);
stm_aix("p5i4","p1i0",[0,"Estado de Resultados Integrales","","",-1,-1,0,"Eri.asp"],165,0);
stm_ep();
stm_aix("p0i3","p0i0",[0,"Salir","","",-1,-1,0,"logoffce.asp","_self","","",ruta+"gnome_logout.png","gnome_logout.png"],100,30);
stm_ep();
stm_em();
