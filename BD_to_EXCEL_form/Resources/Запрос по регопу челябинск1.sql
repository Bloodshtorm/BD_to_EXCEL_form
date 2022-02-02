
                                drop table if exists temp_tariff_dbf;
                                create temp table temp_tariff_dbf as
                                (select * from public.z_get_tariff(3893));
                                select unified_acc_num "UNIFIED_AC",rpa.acc_num "ACC_NUMBER", own.name "OWNER", gdm.name "MU",
                                (select shortname from b4_fias where aoguid = b4fa.place_guid limit 1)::Char(10) "KINDSITY",
                                SUBSTRING(place_name, 4, 999)::Char(40) "CITY", b4f.shortname "KINDSTREET", b4f.formalname "STREET", 
                                b4fa.HOUSE "HOUSE", letter "LETTER", b4fa.housing "HOUSING", b4fa.BUILDING "BUILDING", 
                                gr.croom_num "CROOM_NUM", b4fa.house_guid "ADR_FIAS", '' as "PRIM", gr.carea "PLOSHAD", area_living_owned as "LIVING_SQ",
                                case 
                                when gr.ownership_type = 10 then 'Частная'
                                when gr.ownership_type = 30 then 'Муниципальная'
                                when gr.ownership_type = 40 then 'Государственная'
                                when gr.ownership_type = 50 then 'Коммерческая'
                                when gr.ownership_type = 60 then 'Смешанная'
                                when gr.ownership_type = 80 then 'Федеральная'
                                when gr.ownership_type = 90 then 'Областная'
                                else 'Не указано'
                                end as "OWNERSHIP", 
                                case when gr.IS_COMMUNAL then 'КОММУНАЛЬНАЯ' else 'ОТДЕЛЬНАЯ' end as "HABIT_TYPE",
                                '' as "PROPIS",'041' "SRV_ID",'' as "SRV_NAME",
                                (1) as "REC_TYPE",
                                (select * from temp_tariff_dbf) "TARIF",
                                '0' "NORM",
                                round(psum.charge_tariff, 2) as "SUMMA",
                                round(psum.RECALC, 2) as "RECALC",
                                round(BASE_TARIFF_DEBT, 2) as "DOLG",
                                round(PENALTY, 2) "PENI",
                                round((PENALTY_PAYMENT + TARIFF_PAYMENT), 2) "OPLATA",
                                round(SALDO_OUT_SERV, 2) "SUMMA_K_OP",
                                '40603810209280004926' as "RS",'Филиал "Центральный" Банка ВТБ (ПАО) в г. Москве' as "BANK",'30101810145250000411' as "KOR",
                                '044525411' as "BIK",'454048' as "INDEX", To_char(rp.cstart, 'MMYY') as "PERIOD_OPL", To_char(rp.cstart + interval '1 month', 'DDMMYYYY') as "OPLATIT_DO",'fondkp174@mail.ru' as "EMAIL"
                                from regop_pers_acc_period_summ psum
                                join regop_period rp on rp.id = psum.period_id
                                join regop_pers_acc rpa on rpa.id = psum.account_id and rpa.state_id = 804
                                join gkh_room gr on gr.id = rpa.room_id
                                join regop_pers_acc_owner own on rpa.acc_owner_id = own.id and rpa.state_id = 804
                                join gkh_reality_object ro on ro.id = gr.ro_id
                                join gkh_dict_municipality gdm on ro.municipality_id = gdm.id
                                join b4_fias_address b4fa on b4fa.id = ro.fias_address_id
                                join b4_fias b4f on b4f.aoguid = b4fa.street_guid and b4f.actstatus = 1
                                where psum.period_id = 1608
                                limit 10