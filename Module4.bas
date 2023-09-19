Attribute VB_Name = "Module4"
CREATE OR REPLACE PROCEDURE SIAM.P_SeguimientoGuardaDatosbk1(piSeg IN integer, piSegAnt in integer, piAna IN NUMBER, psFecha IN VARCHAR2, piAct in smallint, piTar in smallint, pires in smallint, piUsi IN smallint, pides smallint, psObs IN VARCHAR2, psDoctos IN VARCHAR2, psProg IN VARCHAR2, piConAcu in integer, psSancion in varchar2, psCondonacion in varchar2, psMemo in varchar2, pcur OUT sys_refcursor)
IS

v_id seguimiento.id%TYPE:=0;

sid_av VARCHAR2(200):='';
sid_doc VARCHAR2(250):='';
scons_fi VARCHAR2(150):='';
sid_fi VARCHAR2(150):='';

s_Oficio seguimientosanción.OFICIO%type:='';
s_infraccion seguimientosanción.FECHAINFRACCIÓN%type:=null;
s_unidad seguimientosanción.iduni%type:=0;
s_monto seguimientosanción.monto%type:=0;
s_importe seguimientosanción.importepesos%type:=0;

s_fecha seguimientocondonación.FECHA%type:=null;
s_porcentaje seguimientocondonación.PORCENTAJE%type:=0;

s_sumamonto number:=0;
s_sumaimp number:=0;
s_procede SEGUIMIENTOSANXANACAU.PROCEDE%type:=0;
s_ana SEGUIMIENTOSANXANACAU.IDANACAU%type:=0;
s_sub SEGUIMIENTOSANXANACAU.sub%type:='';

s_memo seguimientomemos.MEMORANDO%type:='';
d_fechamemo seguimientomemos.FECHA%type:=null;
--s_Oficio SANCIONES.NO_OFICIO_RESOLUCIÓN%TYPE:='';
--s_Resolucion SANCIONES.F_OFICIO_RESOLUCIÓN%TYPE:=NULL;
--s_Infraccion SANCIONES.F_INFRACCIÓN%TYPE:=NULL;
--s_Dias SANCIONES.DÍAS_MULTA%TYPE:=0;
--s_Importe SANCIONES.IMP_MULTA%TYPE:=0;
--s_NoMulta SANCIONES.NO_MULTA%TYPE:='';

i1 INTEGER:=0;
i2 INTEGER:=0;
i3 INTEGER:=0;
ss VARCHAR2(500):='';
sErr VARCHAR2(200):='';
sSen VARCHAR2(4000):='';
s1 VARCHAR2(200):='|';
s VARCHAR2(500):='';
dFecha DATE:=NULL;
dFecha1 DATE:=TO_DATE(psFecha,'dd/mm/yyyy hh24:mi:ss');
i INTEGER:=F_Idconsecutivo('id','seguimiento');


CURSOR x_cur IS SELECT id, idusi FROM seguimiento WHERE id BETWEEN i-2 AND i-1 FOR UPDATE;
BEGIN
   COMMIT;
   Execute immediate 'alter session set nls_numeric_characters=''.,''';
   IF piSeg>0 THEN --Se trata de actualización de los datos del Seguimiento

       v_id:=piSeg;

       UPDATE seguimiento SET idtar=piTar,idusi=piUsi,fecha=dFecha1,actualización=SYSDATE,iddes=CASE WHEN piDes>0 THEN piDes ELSE NULL END,idres=piRes
       WHERE id=piSeg;

   Else

       FOR x_rec IN x_cur
       Loop
           UPDATE seguimiento SET idusi=idusi WHERE CURRENT OF x_cur;
       END LOOP;
       
       v_id:=F_Idconsecutivo('id','seguimiento');
       INSERT INTO seguimiento (id,idant,idana,idact,idtar,idres,idusi,fecha,iddes) VALUES
       (v_id,CASE WHEN piSegAnt<=0 THEN v_id ELSE piSegAnt END,piAna,piAct,piTar,piRes,piUsi,dFecha1,CASE WHEN piDes>0 THEN piDes ELSE NULL END);
       
       select count(*) into i from seguimientoprog where idant=piSegAnt and idact=piAct;
       if i>0 then --Actualiza idseg de la tabla de actividadprogramada (seguimientoprog)
             update seguimientoprog set idseg=v_id where idant=piSegAnt and idact=piAct;
       end if;

   END IF;
   IF LENGTH(psObs)>5 THEN ---guarda Observaciones
          if piseg>0 then --Actualización
              select count(*) into i1 from seguimientoobs where idseg=v_id;
       Else
              i1:=0;
       end if;
       If i1 > 0 Then
              update seguimientoobs set observaciones=psObs where idseg=v_id;
       Else
              INSERT INTO seguimientoobs (idseg,observaciones) VALUES
              (v_id,psObs);
       end if;
   elsif piseg>0 then --Actualización
       delete from seguimientoobs where idseg=v_id;
   END IF;

   s1:=psDoctos; --Documentos
   ss:=psProg; --Acts programadas

   IF SUBSTR(s1,LENGTH(s1),1)<>'|' THEN
         s1:=s1||'|';
   END IF;
   s:='';
   WHILE INSTR(s1,'|')>0 --Guarda los documentos del avance, Se obtiene un valor por cada documento
   Loop
          i2:=F_Obtienecamponumero(s1); --obtiene el iddoc
       if piSeg>0 then --Actualización
              select count(*) into i1 from seguimientodoctos where idseg=v_id and iddoc=i2;
       Else
              i1:=0;
       end if;
       if i1>0 then -- actualiza
              update seguimientodoctos set actualización=sysdate where idseg=v_id and iddoc=i2;
       Else
              INSERT INTO seguimientodoctos (id,idseg,iddoc,registro,actualización) VALUES (f_idconsecutivo('id','seguimientodoctos'),v_id,i2,sysdate,sysdate);
       end if;
       s:=s||i2||',';
   END LOOP;
   if piSeg>0 then --Actualización
       --Elimina las documentos no emidtidos en caso que se trate de Actualización
       delete from seguimientodoctos where idseg=v_id and instr(','||s,','||iddoc||',')=0;
   end if;

   IF SUBSTR(ss,LENGTH(ss),1)<>'|' THEN
         ss:=ss||'|';
   END IF;
   s:='';
   WHILE INSTR(ss,'|')>0 --Se espera 3 valores por cada subcadena de Actividad Programada
   Loop
          i2:=F_Obtienecamponumero(ss); --idact
          dFecha:=F_Obtienecampofecha(ss); --FechaProgramada
          i3:=F_Obtienecamponumero(ss); --idusi
       if piSeg>0 then --Actualización
              select count(*) into i1 from seguimientoprog where idant=v_id and idact=i2;
       Else
              i1:=0;
       end if;
       If i1 > 0 Then
              update seguimientoprog set idusi=i3,fecha=dfecha where idant=v_id and idact=i2;
       Else
              INSERT INTO seguimientoprog (idant,idana,idact,idusi,fecha) VALUES (v_id,piAna,i2,i3,dFecha);
       end if;
       s:=s||i2||',';
   END LOOP;
   if piSeg>0 then --Actualización
       --Elimina las actividades no programadas en caso que se trate de Actualización
       delete from seguimientoprog where idant=v_id and instr(','||s,','||idact||',')=0;
   end if;
   
   --Otros Datos
   --No. Acuerdo
   i:=0;
   select count(*) into i from seguimientoacuerdos where idseg=v_id;
   If piconacu > 0 Then
          s_oficio:=f_nuevofolio(6,0,piseg);
       if instr(s_oficio,'???')>0 then
              s_oficio:=replace(s_oficio,'???',''||piconacu);
       end if;
           if i>0 then-- se actualiza
               update seguimientoacuerdos set acuerdo=s_oficio where idseg=v_id;
        Else
               insert into seguimientoacuerdos (idseg,acuerdo,registro,año,consecutivo) values (v_id,s_oficio,sysdate,f_folioanio(6,s_oficio),f_folioconsecutivo(6,s_oficio));
        end if;
   Else
          If i > 0 Then
              delete seguimientoacuerdos where idseg=v_id;
       end if;
   end if;
   --No. y fecha Memorando
   i:=0;
   select count(*) into i from seguimientomemos where idseg=v_id;
   s:=psMemo;
   if instr(s,'|')>0 then
           s_memo:=f_obtienecampotexto(s);
           d_fechamemo:=f_obtienecampofecha(s);
           if i>0 then-- se actualiza
               update seguimientomemos set memorando=s_memo,fecha=d_fechamemo where idseg=v_id;
        Else
               insert into seguimientomemos (idseg,memorando,fecha,registro) values (v_id,s_memo,d_fechamemo,sysdate);
        end if;
   Else
          If i > 0 Then
              delete seguimientomemos where idseg=v_id;
       end if;
   end if;
   --Sanción
   i:=0;
   select count(*) into i from seguimientosanción where idseg=v_id;
   s:=psSancion;
   if instr(s,'|')>0 then
          s_oficio:=f_obtienecampotexto(s);
          s_infraccion:=f_obtienecampofecha(s);
          delete from seguimientosanxanacau where idseg=v_id;
          while instr(s,'|')>0
          Loop
              dbms_output.put_line('Cadena sanción X Pro: '||s);
              s_ana:=f_obtienecamponumero(s);
              s_sub:=f_obtienecampotexto(s);
              s_procede:=f_obtienecamponumero(s);
              If s_procede = 0 Then
                  ss:=f_obtienecampotexto(s);
                  ss:=f_obtienecampotexto(s);
                  ss:=f_obtienecampotexto(s);
                  s_unidad:=null;
                  s_monto:=0;
                  s_importe:=0;
              Else
                  s_unidad:=f_obtienecamponumero(s);
                  s_monto:=f_obtienecamponumero(s);
                  s_importe:=f_obtienecamponumero(s);
                  s_sumamonto:=s_sumamonto+s_monto;
                  s_sumaimp:=s_sumaimp+s_monto;
              end if;
              insert into seguimientosanxanacau (idseg,idanacau,sub,procede,fecha,iduni,monto,importepesos,registro) values
              (v_id,s_ana,s_sub,s_procede,dfecha1,s_unidad,s_monto,s_importe,sysdate);
          end loop;
          if i>0 then-- se actualiza
              update seguimientosanción set oficio=s_oficio,fechainfracción=s_infraccion,iduni=s_unidad,monto=s_sumamonto,importepesos=s_sumaimp where idseg=v_id;
     
          Else
              --s_oficio:=f_nuevofolio(4,0,piana);
           --if instr(s_oficio,'???')>0 then
           --       s_oficio:=replace(s_oficio,'???',''||piconsan);
           --end if;
               insert into seguimientosanción (idseg,año,consecutivo,oficio,fechainfracción,iduni,monto,importepesos,registro) values (v_id,f_folioanio(4,s_Oficio),f_folioconsecutivo(4,s_Oficio),s_oficio,s_infraccion,s_unidad,s_sumamonto,s_sumaimp,sysdate);
          end if;
   Else
       If i > 0 Then
              delete seguimientosanción where idseg=v_id;
       end if;
   end if;
   i:=0;
   select count(*) into i from seguimientocondonación where idseg=v_id;
   s:=psCondonacion;
   if instr(s,'|')>0 then
          s_oficio:=f_obtienecampotexto(s);
          s_fecha:=f_obtienecampofecha(s);
          s_porcentaje:=f_obtienecamponumero(s);
          if i>0 then-- se actualiza
           update seguimientocondonación set oficio=s_oficio,fecha=s_fecha,porcentaje=s_porcentaje where idseg=v_id;
       Else
              --s_oficio:=f_nuevofolio(4,0,piana);
           --if instr(s_oficio,'???')>0 then
           --       s_oficio:=replace(s_oficio,'???',''||piconsan);
           --end if;
           insert into seguimientocondonación (idseg,año,consecutivo,oficio,fecha,registro,porcentaje) values (v_id,f_folioanio(5,s_Oficio),f_folioconsecutivo(5,s_Oficio),s_oficio,s_fecha,sysdate,s_porcentaje);
       end if;
   Else
       If i > 0 Then
              delete seguimientocondonación where idseg=v_id;
       end if;
   end if;
 COMMIT;
    OPEN pcur FOR SELECT v_id as id,'Folio' as folio FROM dual;
          
    EXCEPTION
     WHEN NO_DATA_FOUND THEN
           ROLLBACK;
           sErr:='SQLCODE: '||SQLCODE||'  SQLERRM: '||SQLERRM;
          sSen:=piSeg||'¬'||piSegAnt||'¬'||piAna||'¬'||psFecha||'¬'||piAct||'¬'||piTar||'¬'||piUsi||'¬'||psObs||'¬'||psDoctos||'¬'||psProg||'¬'||piConAcu||'¬'||psSancion||'¬'||psCondonacion||'¬'||psMemo;
          sSen:=sSen||SUBSTR(sErr,1,2000-LENGTH(sSen));
          INSERT INTO ERR_SEGUIMIENTO (cadena,tipo) VALUES (sSen,1);
          COMMIT;
           OPEN pcur FOR SELECT 0 as id,'Datos no encontrados '||sErr AS Folio FROM dual;
     WHEN OTHERS THEN
           ROLLBACK;
           sErr:='SQLCODE: '||SQLCODE||'  SQLERRM: '||SQLERRM;
          sSen:=piSeg||'¬'||piSegAnt||'¬'||piAna||'¬'||psFecha||'¬'||piAct||'¬'||piTar||'¬'||piUsi||'¬'||psObs||'¬'||psDoctos||'¬'||psProg||'¬'||piConAcu||'¬'||psSancion||'¬'||psCondonacion||'¬'||psMemo;
          sSen:=sSen||SUBSTR(sErr,1,2000-LENGTH(sSen));
          INSERT INTO ERR_SEGUIMIENTO (cadena,tipo) VALUES (sSen,2);
          COMMIT;
           OPEN pcur FOR SELECT -1 as id, 'Error no esperado'||sErr AS Folio FROM dual;

          dbms_output.put_line(sErr);
     RAISE;
END P_SeguimientoGuardaDatosbk1;
/

