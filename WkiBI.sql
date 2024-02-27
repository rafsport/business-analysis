--- Anagrafica Anna

Select * From dbo.AnagraficaAnna_FromFS;
select count (distinct Tenant) as Tenant , count (distinct AAATenantId) as AAATenantId from AnagraficaAnna_FromFS; -- 870 (credo prenda solo i Tenant di Genya)



Select * From dbo.AnagraficaAnna_FromGY;
select count (distinct Tenant) as Tenant , count (distinct AAATenantId) as AAATenantId from AnagraficaAnna_FromGY; -- 6.569 AAATenantid (esistono NULL), 6.585 Tenant

Select * From dbo.AnagraficaAnna_FromWD;
select count (distinct OfficeId) as OfficeId , count (distinct CodOffice) as AAATenantId from AnagraficaAnna_FromWD; -- 15.662


--- Anagrafica Ben

Select * From dbo.AnagraficaBen_FromFS;
select
	count (*),
	count (distinct BenId) as BenId ,
	count (distinct CodiceFiscale) as CodiceFiscale,
	count (distinct PartitaIVA) as PartitaIva,
	count (distinct BenWebdeskId) as BenWebdeskId, -- Webdesk
	count (distinct AaaBenId) as AaaBenId -- Genya
from AnagraficaAnna_FromFS;


Select * From dbo.AnagraficaBen_FromGY;
Select
	count (*),
	count (distinct AaaBenId) as AaaBenId,
	count (distinct cast(A.Tenant as varchar) + '.' + cast(A.IdSubject as varchar) as BenId),
from AnagraficaAnna_FromGY;


-- Conteggio dei Ben abilitati da Anna (solo di Genya per ora) all'acquisto da eShop che hanno anche fatto 'AskAnna'
Select
	Totale=count(*)
From
	AnagraficaBen B
	Left Join HookUsageData H On B.BenId = H.BenId
Where
	EnableSkillEShop = 1 And H.Azione = 'AskAnna'		-- 1453 abilitati all'acquisto,	17390 hanno fatto AskAnna --> solo 20



-- Stesso tracciato di AnagraficaBen_FromFS con una colonna in più EnableSkillEShop che riporta - SOLO PER GENYA - lo stato del 'flaggino' o del 'flaggone' messi in or.
Select
	F.*,
	EnableSkillEShop = Convert(bit, Case When G.EnableSkillEShop4Ben = 1 or A.EnableSkillEShop4Tenant = 1 then 1 else 0 end),
	G.EnableSkillEShop4Ben,
	A.EnableSkillEShop4Tenant
From
	          AnagraficaBen_FromFS  F
	left join AnagraficaBen_FromGY  G On F.AaaTenantId = G.AaaTenantId And F.AaaBenId = G.AaaBenId
	left join AnagraficaAnna_FromGY A On A.AaaTenantId = G.AaaTenantId
