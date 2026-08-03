# -*- coding: utf-8 -*-
"""
Industry seed packs for the D365 F&SC Payment Prediction demo generator.

A seed pack is only used ONCE -- to create the Excel workbook the first time.
After that, the workbook's `Input_Customers` sheet is the source of truth and
you edit it directly for each new demo. See README.md.

To add a new industry: copy AIRPORT_SBN, change the customers + BILLING_LINES
entries for the groups you use, and register it in SEED_PACKS at the bottom.

Columns in each customer tuple-dict:
  account   Customer account number shown in the D365 grid
  name      Customer name
  group     Customer group  (drives billing lines AND is a prediction "top factor")
  terms     Payment terms   (Net 15 / Net 30 / Net 45 / Net 60)
  risk      0.0 - 1.0 payment-reliability score. 1.0 = always pays on time.
            This is the single biggest driver of the on-time probability.
  aging     current | mild | moderate | severe  -- how past-due their open items are
  opens     how many open transactions to generate
  city      short location string used in the grid tooltip
  address   multi-line address for the "Customer details" FactBox
"""

# ---------------------------------------------------------------------------
# Billing lines per customer group.
# (description, transaction_type, amount_low, amount_high)
# transaction_type must be one of:
#   Customer | Collection letter | Project | Sales order | Interest note
# "Customer" is what a posted free text invoice shows as in the D365 grid.
# ---------------------------------------------------------------------------
BILLING_LINES = {
    "Airlines - Signatory": [
        ("Landing fees - signatory rate", "Customer", 9_400, 27_500),
        ("Terminal rent - ticket counter & gate", "Customer", 16_800, 41_200),
        ("Passenger boarding bridge use fee", "Customer", 3_100, 8_900),
        ("Common use terminal equipment (CUTE)", "Customer", 2_400, 6_700),
    ],
    "Airlines - Regional": [
        ("Landing fees - non-signatory rate", "Customer", 4_200, 13_800),
        ("Terminal rent - shared use", "Customer", 6_500, 18_400),
        ("Ramp & RON parking fees", "Customer", 1_800, 5_400),
    ],
    "Cargo & Freight": [
        ("Cargo landing fees", "Customer", 5_600, 19_300),
        ("Cargo apron & building lease", "Customer", 11_200, 28_700),
        ("Fuel flowage fee - cargo", "Customer", 1_400, 4_800),
    ],
    "Rental Car Concession": [
        ("Concession fee - MAG true-up", "Customer", 24_000, 118_000),
        ("Counter & ready-return space rent", "Customer", 8_900, 26_400),
        ("Customer facility charge remittance", "Customer", 12_300, 44_600),
    ],
    "Terminal Concessions": [
        ("Concession fee - percentage of gross", "Customer", 4_100, 21_800),
        ("Terminal space rent", "Customer", 3_600, 14_200),
        ("Utilities & CAM recovery", "Customer", 900, 4_300),
    ],
    "Advertising": [
        ("Advertising concession - minimum annual guarantee", "Customer", 5_200, 16_900),
        ("Digital display placement fee", "Customer", 1_600, 6_100),
    ],
    "FBO & Fuel": [
        ("FBO ground lease", "Customer", 12_600, 38_500),
        ("Fuel flowage fee", "Sales order", 2_800, 11_400),
        ("Into-plane servicing permit", "Customer", 1_200, 4_100),
    ],
    "MRO & Hangar": [
        ("Hangar lease - maintenance facility", "Customer", 7_400, 22_900),
        ("Land lease - airside parcel", "Customer", 3_200, 9_800),
        ("Utilities & CAM recovery", "Customer", 700, 3_100),
    ],
    "Ground Handling": [
        ("Ground handling permit fee", "Customer", 2_600, 8_700),
        ("Ramp access & badging recovery", "Customer", 800, 3_400),
    ],
    "Parking & Shuttle": [
        ("Parking concession fee", "Customer", 38_000, 142_000),
        ("Shuttle operations management fee", "Customer", 9_400, 24_800),
    ],
    "Hangar & Land Lease": [
        ("T-hangar rent", "Customer", 420, 1_850),
        ("Corporate hangar lease", "Customer", 2_900, 11_600),
        ("Land lease - general aviation parcel", "Customer", 780, 4_200),
        ("Tie-down & apron parking", "Customer", 180, 940),
    ],
    "Charter Operator": [
        ("Commercial operating permit", "Customer", 1_400, 5_600),
        ("Landing fees - charter", "Customer", 900, 4_300),
        ("Hangar rent - charter fleet", "Customer", 2_200, 8_100),
    ],
    "Government": [
        ("Terminal space lease - federal inspection/office", "Customer", 6_800, 23_400),
        ("AIP grant reimbursable - runway rehabilitation", "Project", 84_000, 460_000),
        ("Reimbursable services agreement", "Project", 14_500, 68_000),
    ],
    "Education & Charter": [
        ("Charter operations & apron fee", "Customer", 3_400, 14_700),
        ("Hangar rent - team aircraft", "Customer", 5_100, 17_300),
        ("Aviation program facility lease", "Customer", 1_900, 7_400),
    ],

    # --- Benefit fund administration ---------------------------------------
    "Taft-Hartley Health & Welfare": [
        ("Administration fee - per participant per month", "Customer", 48_000, 182_000),
        ("Claims processing & adjudication services", "Customer", 22_000, 86_000),
        ("Eligibility & enrollment administration", "Customer", 9_400, 34_000),
        ("COBRA administration & premium billing", "Customer", 3_200, 12_800),
    ],
    "Taft-Hartley Benefit Trust": [
        ("Fund administration fee - per participant per month", "Customer", 36_000, 145_000),
        ("Pension benefit calculation & processing", "Customer", 14_000, 52_000),
        ("Trustee meeting & fund accounting support", "Customer", 6_800, 24_500),
        ("Form 5500 preparation & audit support", "Project", 18_000, 72_000),
    ],
    "Contributing Employer": [
        ("Employer contribution remittance - monthly", "Customer", 28_000, 164_000),
        ("Delinquency assessment & liquidated damages", "Interest note", 1_400, 9_600),
        ("Payroll compliance audit recovery", "Project", 12_000, 68_000),
        ("ACA reporting & Form 1095-C services", "Customer", 2_600, 9_800),
    ],
    "Contributing Employer - Construction": [
        ("Employer contribution remittance - certified payroll", "Customer", 34_000, 196_000),
        ("Payroll compliance audit recovery", "Project", 18_000, 94_000),
        ("Delinquency assessment & liquidated damages", "Interest note", 2_200, 14_500),
        ("Wage & fringe reconciliation services", "Customer", 4_800, 17_200),
    ],
    "Public Sector Benefit Plan": [
        ("Plan administration fee - per employee per month", "Customer", 42_000, 158_000),
        ("Retiree & OPEB benefit administration", "Customer", 16_500, 62_000),
        ("Open enrollment services & member communications", "Customer", 7_200, 28_400),
        ("Actuarial valuation & GASB 75 reporting support", "Project", 24_000, 96_000),
    ],
}


# ---------------------------------------------------------------------------
# South Bend International Airport (SBN) / St. Joseph County Airport Authority
# ---------------------------------------------------------------------------
# NOTE ON NAMES: real tenant/airline brands are used because recognising their
# own AR ledger is what makes the demo land. All balances, dates, probabilities
# and payment behaviour are SYNTHETIC. Accounts with genuinely poor payment
# behaviour are deliberately fictional entities, so no real organisation is
# portrayed as delinquent. Keep that convention if you edit the list.

AIRPORT_SBN = {
    "config": {
        "Demo name": "South Bend International Airport",
        "Legal entity": "SBN",
        "Organization": "St. Joseph County Airport Authority",
        "User name": "Demo User",
        "User initials": "",
        "User email": "demo.user@contoso.com",
        "Currency": "USD",
        "As of date": "2026-07-23",
        "Aging period definition": "30_60_90_180",
        "Red dot threshold (%)": 60,
        "Theme": "red",
        "Size": "comfortable",
        "Random seed": 20260723,
    },
    "customers": [
        # --- Airlines ------------------------------------------------------
        dict(account="AIR-1001", name="Allegiant Air, LLC", group="Airlines - Signatory",
             terms="Net 30", risk=0.88, aging="mild", opens=4,
             city="Las Vegas, NV",
             address="1201 N Town Center Drive\nLas Vegas, NV 89144\nUSA"),
        dict(account="AIR-1002", name="Delta Air Lines, Inc.", group="Airlines - Signatory",
             terms="Net 30", risk=0.92, aging="current", opens=3,
             city="Atlanta, GA",
             address="1030 Delta Boulevard\nAtlanta, GA 30354\nUSA"),
        dict(account="AIR-1003", name="United Airlines, Inc.", group="Airlines - Signatory",
             terms="Net 30", risk=0.90, aging="current", opens=3,
             city="Chicago, IL",
             address="233 S Wacker Drive\nChicago, IL 60606\nUSA"),
        dict(account="AIR-1004", name="American Airlines, Inc.", group="Airlines - Signatory",
             terms="Net 30", risk=0.89, aging="mild", opens=3,
             city="Fort Worth, TX",
             address="1 Skyview Drive\nFort Worth, TX 76155\nUSA"),
        dict(account="AIR-1005", name="SkyWest Airlines, Inc.", group="Airlines - Regional",
             terms="Net 30", risk=0.84, aging="mild", opens=3,
             city="St. George, UT",
             address="444 South River Road\nSt. George, UT 84790\nUSA"),
        dict(account="AIR-1006", name="Republic Airways Inc.", group="Airlines - Regional",
             terms="Net 30", risk=0.81, aging="moderate", opens=3,
             city="Indianapolis, IN",
             address="8909 Purdue Road, Suite 300\nIndianapolis, IN 46268\nUSA"),
        dict(account="AIR-1007", name="Air Wisconsin Airlines LLC", group="Airlines - Regional",
             terms="Net 30", risk=0.78, aging="moderate", opens=3,
             city="Appleton, WI",
             address="W6390 Challenger Drive, Suite 203\nAppleton, WI 54914\nUSA"),

        # --- Cargo ---------------------------------------------------------
        dict(account="CGO-1201", name="FedEx Corporate Services, Inc.", group="Cargo & Freight",
             terms="Net 30", risk=0.93, aging="current", opens=3,
             city="Memphis, TN",
             address="3620 Hacks Cross Road\nMemphis, TN 38125\nUSA"),
        dict(account="CGO-1202", name="United Parcel Service Co.", group="Cargo & Freight",
             terms="Net 30", risk=0.91, aging="current", opens=2,
             city="Atlanta, GA",
             address="55 Glenlake Parkway NE\nAtlanta, GA 30328\nUSA"),
        dict(account="CGO-1203", name="Mountain Air Cargo, Inc.", group="Cargo & Freight",
             terms="Net 30", risk=0.74, aging="mild", opens=3,
             city="Kannapolis, NC",
             address="3524 Airport Road\nKannapolis, NC 28081\nUSA"),

        # --- Rental car ----------------------------------------------------
        dict(account="RAC-2001", name="Avis Budget Car Rental, LLC", group="Rental Car Concession",
             terms="Net 30", risk=0.86, aging="mild", opens=3,
             city="Parsippany, NJ",
             address="6 Sylvan Way\nParsippany, NJ 07054\nUSA"),
        dict(account="RAC-2002", name="The Hertz Corporation", group="Rental Car Concession",
             terms="Net 30", risk=0.72, aging="moderate", opens=4,
             city="Estero, FL",
             address="8501 Williams Road\nEstero, FL 33928\nUSA"),
        dict(account="RAC-2003", name="Enterprise Leasing Company of Indianapolis, LLC",
             group="Rental Car Concession",
             terms="Net 30", risk=0.88, aging="current", opens=3,
             city="Indianapolis, IN",
             address="8425 Zionsville Road\nIndianapolis, IN 46268\nUSA"),
        dict(account="RAC-2004", name="Sixt Rent A Car, LLC", group="Rental Car Concession",
             terms="Net 30", risk=0.68, aging="moderate", opens=3,
             city="Fort Lauderdale, FL",
             address="1501 Broken Sound Parkway NW\nFort Lauderdale, FL 33487\nUSA"),

        # --- Terminal concessions & advertising ----------------------------
        dict(account="CON-3001", name="SSP America, Inc.", group="Terminal Concessions",
             terms="Net 30", risk=0.76, aging="mild", opens=3,
             city="Ashburn, VA",
             address="20408 Bashan Drive, Suite 300\nAshburn, VA 20147\nUSA"),
        dict(account="CON-3002", name="Paradies Lagardere", group="Terminal Concessions",
             terms="Net 30", risk=0.74, aging="mild", opens=3,
             city="Atlanta, GA",
             address="2849 Paces Ferry Road SE, Suite 400\nAtlanta, GA 30339\nUSA"),
        dict(account="CON-3003", name="Clear Channel Airports", group="Advertising",
             terms="Net 45", risk=0.70, aging="moderate", opens=2,
             city="Allentown, PA",
             address="5850 Hamilton Boulevard\nAllentown, PA 18106\nUSA"),
        dict(account="CON-3004", name="Michiana Coffee Roasters LLC", group="Terminal Concessions",
             terms="Net 30", risk=0.48, aging="moderate", opens=3,
             city="South Bend, IN",
             address="812 Lincoln Way West\nSouth Bend, IN 46616\nUSA"),
        dict(account="CON-3005", name="Bendix Newsstand & Gifts LLC", group="Terminal Concessions",
             terms="Net 30", risk=0.36, aging="severe", opens=4,
             city="South Bend, IN",
             address="4477 Progress Drive, Suite B\nSouth Bend, IN 46628\nUSA"),

        # --- FBO / fuel / MRO ----------------------------------------------
        dict(account="FBO-4001", name="Atlantic Aviation SBN, LLC", group="FBO & Fuel",
             terms="Net 15", risk=0.85, aging="mild", opens=3,
             city="South Bend, IN",
             address="1741 Lawrence D Bell Drive\nSouth Bend, IN 46628\nUSA"),
        dict(account="FBO-4002", name="Corporate Wings, LLC", group="FBO & Fuel",
             terms="Net 15", risk=0.79, aging="mild", opens=3,
             city="South Bend, IN",
             address="4599 Progress Drive\nSouth Bend, IN 46628\nUSA"),
        dict(account="FBO-4003", name="Avfuel Corporation", group="FBO & Fuel",
             terms="Net 15", risk=0.87, aging="current", opens=2,
             city="Ann Arbor, MI",
             address="47 W Ellsworth Road\nAnn Arbor, MI 48108\nUSA"),
        dict(account="MRO-4101", name="Hoosier Aviation Maintenance, Inc.", group="MRO & Hangar",
             terms="Net 30", risk=0.41, aging="severe", opens=4,
             city="South Bend, IN",
             address="4321 Aeropark Court\nSouth Bend, IN 46628\nUSA"),

        # --- Ground handling / parking -------------------------------------
        dict(account="SVC-4501", name="Unifi Aviation, LLC", group="Ground Handling",
             terms="Net 30", risk=0.73, aging="moderate", opens=3,
             city="Atlanta, GA",
             address="1000 Hartsfield Centre Parkway, Suite 400\nAtlanta, GA 30354\nUSA"),
        dict(account="SVC-4502", name="SP Plus Corporation", group="Parking & Shuttle",
             terms="Net 30", risk=0.80, aging="mild", opens=3,
             city="Chicago, IL",
             address="200 E Randolph Street, Suite 7700\nChicago, IL 60601\nUSA"),
        dict(account="SVC-4503", name="Great Lakes Ground Support LLC", group="Ground Handling",
             terms="Net 30", risk=0.29, aging="severe", opens=4,
             city="Mishawaka, IN",
             address="2915 N Home Street\nMishawaka, IN 46545\nUSA"),

        # --- Hangar / land tenants -----------------------------------------
        dict(account="TEN-5001", name="Bendix Hangar Partners LLC", group="Hangar & Land Lease",
             terms="Net 30", risk=0.55, aging="moderate", opens=3,
             city="South Bend, IN",
             address="4150 Bendix Drive\nSouth Bend, IN 46628\nUSA"),
        dict(account="TEN-5002", name="Kankakee Valley Flyers, Inc.", group="Hangar & Land Lease",
             terms="Net 30", risk=0.33, aging="severe", opens=4,
             city="Walkerton, IN",
             address="9820 W Tyler Road\nWalkerton, IN 46574\nUSA"),
        dict(account="TEN-5003", name="Silver Hawk Aviation LLC", group="Hangar & Land Lease",
             terms="Net 30", risk=0.18, aging="severe", opens=5,
             city="Granger, IN",
             address="13075 Adams Road\nGranger, IN 46530\nUSA"),
        dict(account="TEN-5004", name="Portage Prairie Logistics LLC", group="Hangar & Land Lease",
             terms="Net 30", risk=0.44, aging="moderate", opens=3,
             city="New Carlisle, IN",
             address="31500 Larrison Trail\nNew Carlisle, IN 46552\nUSA"),
        dict(account="TEN-5005", name="Michiana Air Charter LLC", group="Charter Operator",
             terms="Net 15", risk=0.22, aging="severe", opens=4,
             city="South Bend, IN",
             address="4501 Terminal Drive, Hangar 7\nSouth Bend, IN 46628\nUSA"),

        # --- Government / institutional -------------------------------------
        dict(account="GOV-6001", name="Transportation Security Administration",
             group="Government",
             terms="Net 45", risk=0.90, aging="mild", opens=3,
             city="Springfield, VA",
             address="6595 Springfield Center Drive\nSpringfield, VA 20598\nUSA"),
        dict(account="GOV-6002", name="Indiana Department of Transportation", group="Government",
             terms="Net 45", risk=0.86, aging="moderate", opens=3,
             city="Indianapolis, IN",
             address="100 N Senate Avenue, Room N755\nIndianapolis, IN 46204\nUSA"),
        dict(account="GOV-6003", name="St. Joseph County, Indiana", group="Government",
             terms="Net 45", risk=0.88, aging="mild", opens=2,
             city="South Bend, IN",
             address="227 W Jefferson Boulevard\nSouth Bend, IN 46601\nUSA"),
        dict(account="GOV-6004", name="City of South Bend", group="Government",
             terms="Net 45", risk=0.85, aging="mild", opens=2,
             city="South Bend, IN",
             address="227 W Jefferson Boulevard, Suite 1200\nSouth Bend, IN 46601\nUSA"),
        dict(account="EDU-6101", name="University of Notre Dame", group="Education & Charter",
             terms="Net 30", risk=0.91, aging="current", opens=3,
             city="Notre Dame, IN",
             address="724 Grace Hall\nNotre Dame, IN 46556\nUSA"),
        dict(account="EDU-6102", name="Ivy Tech Community College", group="Education & Charter",
             terms="Net 30", risk=0.82, aging="mild", opens=2,
             city="South Bend, IN",
             address="220 Dean Johnson Boulevard\nSouth Bend, IN 46601\nUSA"),
    ],
}


# ---------------------------------------------------------------------------
# Benefit fund administration -- multiemployer (Taft-Hartley) welfare funds and
# benefit trusts, the employers that remit contributions to them, and public
# sector benefit plans.
# ---------------------------------------------------------------------------
# NOTE ON NAMES: every entity in this pack is fictional. The delinquency in the
# book sits on the contributing-employer accounts, which is where it sits in a
# real fund-administration AR ledger -- contribution delinquency is the whole
# collections story in this industry. Public sector plans reference real
# jurisdictions, so they are deliberately kept as reliable payers. Keep that
# convention if you edit the list.

BENEFIT_FUNDS = {
    "config": {
        "Demo name": "Benefit Fund Administration",
        "Legal entity": "USMF",
        "Organization": "Benefit Fund Administration Services",
        "User name": "Demo User",
        "User initials": "",
        "User email": "demo.user@contoso.com",
        "Currency": "USD",
        "As of date": "2026-08-03",
        "Aging period definition": "30_60_90_180",
        "Red dot threshold (%)": 60,
        "Theme": "red",
        "Size": "comfortable",
        "Random seed": 20260803,
    },
    "customers": [
        # --- Taft-Hartley health & welfare funds ----------------------------
        dict(account="C00011", name="Northeast Carpenters Welfare Fund",
             group="Taft-Hartley Health & Welfare",
             terms="Net 30", risk=0.91, aging="current", opens=3,
             city="New York, NY",
             address="395 Hudson Street, 9th Floor\nNew York, NY 10014\nUSA"),
        dict(account="C00014", name="Tri-State Sheet Metal Workers Health Fund",
             group="Taft-Hartley Health & Welfare",
             terms="Net 30", risk=0.86, aging="mild", opens=3,
             city="Newark, NJ",
             address="1180 Raymond Boulevard, Suite 400\nNewark, NJ 07102\nUSA"),
        dict(account="C00015", name="Metropolitan Transit Employees Health Fund",
             group="Taft-Hartley Health & Welfare",
             terms="Net 30", risk=0.89, aging="current", opens=3,
             city="Brooklyn, NY",
             address="175 Remsen Street, 6th Floor\nBrooklyn, NY 11201\nUSA"),
        dict(account="C00016", name="Hudson Valley Laborers Welfare Fund",
             group="Taft-Hartley Health & Welfare",
             terms="Net 30", risk=0.83, aging="mild", opens=3,
             city="Poughkeepsie, NY",
             address="2 Washington Street, Suite 210\nPoughkeepsie, NY 12601\nUSA"),
        dict(account="C00018", name="Greater New York Teamsters Health Fund",
             group="Taft-Hartley Health & Welfare",
             terms="Net 30", risk=0.87, aging="mild", opens=4,
             city="Long Island City, NY",
             address="27-08 42nd Road\nLong Island City, NY 11101\nUSA"),
        dict(account="C00020", name="Empire Plumbers & Pipefitters Health Fund",
             group="Taft-Hartley Health & Welfare",
             terms="Net 30", risk=0.80, aging="mild", opens=3,
             city="Albany, NY",
             address="890 Third Street, Suite 120\nAlbany, NY 12206\nUSA"),
        dict(account="US-105", name="New York Construction Workers Health Fund",
             group="Taft-Hartley Health & Welfare",
             terms="Net 30", risk=0.78, aging="mild", opens=4,
             city="New York, NY",
             address="266 West 37th Street, 14th Floor\nNew York, NY 10018\nUSA"),

        # --- Taft-Hartley benefit trusts ------------------------------------
        dict(account="C00012", name="Empire State Teachers Benefit Trust",
             group="Taft-Hartley Benefit Trust",
             terms="Net 45", risk=0.90, aging="current", opens=3,
             city="Albany, NY",
             address="20 Corporate Woods Boulevard\nAlbany, NY 12211\nUSA"),
        dict(account="C00013", name="Empire State Electrical Workers Benefit Trust",
             group="Taft-Hartley Benefit Trust",
             terms="Net 30", risk=0.85, aging="current", opens=3,
             city="Flushing, NY",
             address="158-11 Harry Van Arsdale Jr. Avenue\nFlushing, NY 11365\nUSA"),
        dict(account="C00017", name="Long Island Operating Engineers Benefit Fund",
             group="Taft-Hartley Benefit Trust",
             terms="Net 30", risk=0.76, aging="moderate", opens=3,
             city="Hauppauge, NY",
             address="600 Motor Parkway, Suite 305\nHauppauge, NY 11788\nUSA"),
        dict(account="C00019", name="Atlantic Ironworkers Welfare Trust",
             group="Taft-Hartley Benefit Trust",
             terms="Net 30", risk=0.71, aging="moderate", opens=3,
             city="Jersey City, NJ",
             address="101 Hudson Street, Suite 2100\nJersey City, NJ 07302\nUSA"),

        # --- Contributing employers -----------------------------------------
        dict(account="C00021", name="Hudson Manufacturing Group",
             group="Contributing Employer",
             terms="Net 30", risk=0.68, aging="moderate", opens=3,
             city="Yonkers, NY",
             address="1 Alexander Street, Building C\nYonkers, NY 10701\nUSA"),
        dict(account="C00022", name="Empire Logistics Holdings",
             group="Contributing Employer",
             terms="Net 30", risk=0.52, aging="moderate", opens=4,
             city="Elizabeth, NJ",
             address="1000 Corbin Street\nElizabeth, NJ 07201\nUSA"),
        dict(account="C00023", name="Northeast Food Distribution Inc.",
             group="Contributing Employer",
             terms="Net 30", risk=0.44, aging="severe", opens=4,
             city="Bronx, NY",
             address="355 Food Center Drive, Unit 12\nBronx, NY 10474\nUSA"),
        dict(account="C00026", name="Metro Property Management Group",
             group="Contributing Employer",
             terms="Net 30", risk=0.57, aging="moderate", opens=3,
             city="New York, NY",
             address="230 West 41st Street, Suite 1500\nNew York, NY 10036\nUSA"),
        dict(account="C00028", name="Harbor Industrial Services",
             group="Contributing Employer",
             terms="Net 15", risk=0.26, aging="severe", opens=5,
             city="Bayonne, NJ",
             address="140 East 22nd Street, Terminal 3\nBayonne, NJ 07002\nUSA"),
        dict(account="C00029", name="Pioneer Energy Solutions",
             group="Contributing Employer",
             terms="Net 30", risk=0.42, aging="moderate", opens=4,
             city="Stamford, CT",
             address="750 Washington Boulevard, Suite 800\nStamford, CT 06901\nUSA"),
        dict(account="C00030", name="Northeast Healthcare Services Group",
             group="Contributing Employer",
             terms="Net 30", risk=0.66, aging="mild", opens=3,
             city="White Plains, NY",
             address="445 Hamilton Avenue, Suite 1102\nWhite Plains, NY 10601\nUSA"),

        # --- Contributing employers (construction) ---------------------------
        dict(account="C00024", name="Atlantic Construction Services LLC",
             group="Contributing Employer - Construction",
             terms="Net 15", risk=0.31, aging="severe", opens=5,
             city="Brooklyn, NY",
             address="63 Flushing Avenue, Building 280\nBrooklyn, NY 11205\nUSA"),
        dict(account="C00025", name="Liberty Building Products Corporation",
             group="Contributing Employer - Construction",
             terms="Net 30", risk=0.49, aging="moderate", opens=3,
             city="Paterson, NJ",
             address="1 Cianci Street, Suite 40\nPaterson, NJ 07505\nUSA"),
        dict(account="C00027", name="Keystone Infrastructure Partners",
             group="Contributing Employer - Construction",
             terms="Net 30", risk=0.19, aging="severe", opens=5,
             city="Newark, NJ",
             address="744 Broad Street, Suite 1900\nNewark, NJ 07102\nUSA"),

        # --- Public sector benefit plans -------------------------------------
        dict(account="C00031", name="City of Albany Employee Benefits Trust",
             group="Public Sector Benefit Plan",
             terms="Net 45", risk=0.88, aging="mild", opens=3,
             city="Albany, NY",
             address="24 Eagle Street, Room 102\nAlbany, NY 12207\nUSA"),
        dict(account="C00032", name="Suffolk County Employee Health Plan",
             group="Public Sector Benefit Plan",
             terms="Net 45", risk=0.85, aging="current", opens=3,
             city="Hauppauge, NY",
             address="100 Veterans Memorial Highway\nHauppauge, NY 11788\nUSA"),
        dict(account="C00033", name="New Jersey Municipal Benefits Consortium",
             group="Public Sector Benefit Plan",
             terms="Net 45", risk=0.82, aging="mild", opens=3,
             city="Trenton, NJ",
             address="222 West State Street, Suite 300\nTrenton, NJ 08608\nUSA"),
        dict(account="C00034", name="Massachusetts Public Employees Health Trust",
             group="Public Sector Benefit Plan",
             terms="Net 60", risk=0.86, aging="mild", opens=3,
             city="Boston, MA",
             address="19 Staniford Street, 4th Floor\nBoston, MA 02114\nUSA"),
        dict(account="C00035", name="Metropolitan School District Benefits Program",
             group="Public Sector Benefit Plan",
             terms="Net 45", risk=0.79, aging="mild", opens=3,
             city="Syracuse, NY",
             address="725 Harrison Street, Suite 200\nSyracuse, NY 13210\nUSA"),
    ],
}


SEED_PACKS = {
    "airport": AIRPORT_SBN,
    "benefitfunds": BENEFIT_FUNDS,
}
