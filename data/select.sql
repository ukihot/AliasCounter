--*Datatitle shohin
select
    SH_SHCD,
    SH_SHNM,
    SH_KIKAKU,
    SH_REGID,
    SH_REGYMD
from
    m_shohin
;
--*Datatitle uriage
select
    UM_SHCD,
    UM_SHNM,
    UM_KIKAKU,
    UM_REGID,
    UM_REGYMD
from
    T_URIAGEM
where
    UM_D_REGYMD >= to_date('2015-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss')
or  UM_D_EDTYMD >= to_date('2015-01-01 00:00:00', 'yyyy-mm-dd hh24:mi:ss')
order by
    UM_D_EDTYMD desc
;