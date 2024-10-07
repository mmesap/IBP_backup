string = """
American Petroleum Co. Inc.
Campofresco Corp.
Caribe Federal Credit Union.
CIC Construction Group.
Cooperativa de Ahorro y crédito de Aguadilla
Cooperativa De Ahorro Y Credito De Arecibo - CooPACA
Credicentro Coop
Cabo Rojo Coop.
Cooperativa de Ahorro y crédito de Juana Diaz
Cooperativa de Ahorro y crédito de Manati
Coop Rincón
Vega Coop
COSVI / Cooperativa de Seguros de Vida de Puerto Rico
Empire Gas Co. Inc
Empresas de Gas Inc
Maderera Donestevez
AON Risk PR
Bella Group
Borinquen Container Corp.
Caribe Hydroblasting Corrp.
Columbia Central University
Duarte Waste DBA Los Hermanos Corp
Ferraiuoli LLC
Goya de P.R. Inc
Green Care P.R
Humana
J.F Montalvo Cash & Carry Inc.
Jose Santiago Inc
Martin-Brower P.R Inc
Medicoop
Molinos de PR. LLC
Plaza Provisions Co.
Right Way Inc.
Rimco LLC
Sea Air Systems Inc
SuperMax Supermercados Maximos Inc.
Supermercado Agueybana Inc.
Supermercados Mr.Special
to-Ricos Ltd
USIC Group.
V Architecture PSC
Mech- Tech College LLC
Instituto Tecnologico de PR / Recinto San Juan
Humacao Community College
P.R Convention Center
Cabrera Auto
Central Industrial Services
Muñoz Holdings Inc / Ponce Plaza Hotel and Casino
CBX Global
Rodi Cargo Puerto Rico
Sonnell Truck Center
Camioneros Cooperativa de Transporte de Carga
Transporte Santiago & JSV Logistics Inc
Jose Flores Inc.
Rene Ortiz Villafañe
Conwaste
Clean Harbors Caribe Inc.
A&A Waste Management Inc
Central Waste Services Inc
Toyota de Puerto Rico
Coca-Cola P.R Bottlers
Alpharma Integrated Solution Corp.
Luis A. Ayala Colón Sucrs. Inc
Simed
Fulcro Insurance Inc.
Hospital Damas inc.
Hospital de la concepción
Money House Inc.
Villavicencio & Associates Construction Corp
Ranger American of P.R
Tolomeo Bank International Corp
Popular Insurance
Assurant
Capex Financial Co. INC
Central Credit Corp 7 Coop. Seguros Múltiples
Easeway de P.R Corp. / Alberic Colón
Fast Money Financial LLC
F&R Construction Group Inc.
Bonneville Contracting & Technology Group
González Trading LLC / Teselta
Puerto Rico Wire Group
Empresas Puertoriqueñas de Desarrollo Inc.
B.V Properties Inc
Bermúdez, Longo, Díaz-Masso LLC
Newmark
Reality Realty
Universidad del Sagrado Corazón
Ponce Health Sciences University
Máximo Solar Industries
Universal Solar Products Inc.
Power Group
Aireko Energy Solutions
MMM Holding LLC
Vidal & Rodriguez Insurance Brokers
Municipio de Bayamón
Bayamón Medical Center
Hospital Perea
Metropolitano Dr. Pila
Hospital Metropolitano Dr. Susoni
Centro Médico Wilma Vázquez
P.R Women's & Children's Hospital
Hospital Metropolitano San Germán
Hospital San Carlos Borromeo
Zimmer Manufacturing BV
Mar-Co Industries
Neolpharma Inc. Cagua
Avara Pharmaceutical Services
Driven PSC
Sol Puerto Rico Limited (Sol)
Hera Printing Corp
RSM
RRT Rodriguez, Rivera & toro PSC
Cancio, Nadal & Rivera LLC
Adsuar Muñis Goyco Seda & Perez-Ochoa
Delgado I Fernandez LLC
Verdanza Hotel
PAN PEPIN, INC.
Glasstra Group LLC
Island Stevedoring Inc
Motorambar
Hyundai de P.R
Total Energies
Plaza Food Systems
Campofresco Corp
American Petroleum
Empire Gas Group
CIC Construction Group
Empresa Maderera Donestevez
Cooperativa de Ahorro y crédito de Arecibo
Caribe Federal Credit Union
Cooperativa de Ahorro y crédito de Vega Alta
Credicentro Coop.
Cooperativa de Ahorro y crédito de Manati / Coop Manati
Cooperativa de Ahorro y crédito de Juana Díaz
Cooperativa de Seguros de Vida de PR
Trans-Oceanic Life Insurance (TOLIC)
Universal Insurance
United Surety & Indemnity Co.
(USI) United Surety & Indemnity Co.
United Insurance Finance Co., Inc
EDP University of PR Inc
Molinos de Puerto Rico, Inc. / Ardent Mills
Gargiulo P.R Inc
Landfill Technologies LLC& Affiliates / Conwaste
Environmental Quality Laboratories Inc. (Eqla)
Central Industrial Services, LLC
Caribe Hydroblasting Environmental Division (CHED)
Cesar Castillo Inc
Muñoz Holdings Inc
RR Donnelley de P.R Corp
La Rosa del Monte Express Inc.
Molson Coors Brewing
Congar International Corp.
Manatí Medical Center
General Electric Capital Corp.
International Restaurant Services, Inc
Officemax, Inc.
The Ritz Carlton San Juan Hotel Spa & C.
SPECIAL CARE PHARMACY SERVICES
RALPH''S FOOD WAREHOUSE, INC.
Progresive Sales & Service Inc
P.R Food & Paper Dist. Inc
Holsum de P.R. Inc
Pasteleria Cidrines Inc
Compañia Cervecera de P.R Inc
Campofreco Corp
https://www.linkedin.com/in/magda-irizarry-24b78728/
Sistema de Salud Menonita
Empresas Donestevez
Aireko Construction
Tamrio Inc.
Cooperativa de Ahorro y crédito de Rincón
Cooperativa de Ahorro y crédito de Barranquilla
Cooperativa de Ahorro y crédito de Cabo Rojo
Cooperativa de Ahorro y crédito de San Jose
Drogueria Betances LLC
Grupo Cooperativa de seguros Multiples
First Medical Health
Multinacional Life Insurance
Crédito Familia Financial Services INC
Metro Island Mortgage Inc.
Marsh Saldaña Inc
Carrion, Laffitte & Casellas Inc (HUB International)
J. Jaramillo Insurance
Aon Risk Solutions
RD Capital Group Inc.
ICPR Junior College
Instituto Tecnologico de PR
Yaras Corp
Agro Servicios Inc.
Landfill Technologies LLC& Affiliates
Enviromental Quality Laboratories Inc.
Enviromental Resources Management P.R Inc.
Central Industrial Services 
Right Way Enviromental Contractors
DHL Supply Chain
VWR Part of Avantor
Landscape Contractors & Designers LLC
Twins Landscaping Corp.
Sonnell Truck Center 
Consolidated Waste Services LLC & Affiliates
"""
EMPRESAS = string.split("\n")

for empresa in EMPRESAS:
    print(empresa.upper().strip().replace("  ", " ").replace(",", "").replace(".", "").replace("(", "").replace(")", ""))