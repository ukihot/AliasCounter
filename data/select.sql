--*Datatitle shohin
select
    SH_SHCD,
    SH_SHNM,
    SH_KIKAKU,
    SH_REGID,
    SH_REGYMD
from
    m_shohin
where
    m_shohin.SH_D_REGYMD >= to_date('2015-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss')
or  m_shohin.SH_D_EDTYMD >= to_date('2015-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss')
;
--*Datatitle uriage
select
    T_URIAGEM.UM_SHCD,
    T_URIAGEM.UM_SHNM,
    T_URIAGEM.UM_KIKAKU,
    M_TANTO.TA_TANM,
    T_URIAGEM.UM_KEIYMD,
    M_TOKUI.TK_TKNM,
    T_URIAGEM.UM_SINM
from
    T_URIAGEM
    inner join
        M_TOKUI
    on  T_URIAGEM.UM_TKCD = M_TOKUI.TK_TKCD
    inner join
        M_TANTO
    on  M_TOKUI.TK_UTACD = M_TANTO.TA_TACD
where
    T_URIAGEM.UM_D_REGYMD >= to_date('2015-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss')
or  T_URIAGEM.UM_D_EDTYMD >= to_date('2015-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss')
order by
    T_URIAGEM.UM_KEIYMD desc
;
