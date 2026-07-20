# Pending mapping decisions

Open items that need a stakeholder answer before the migration is fully correct.
Each has a safe placeholder in place today so the migration runs without
failing — update the relevant function in `ProfileMigration/Program.cs` once
an answer comes back, then remove the entry below.

---

## ~~PROFILES_TB.MARITAL_STATUS_ID — old value 5 ("خاطب"/Engaged)~~ — resolved

- **Resolution**: stakeholder mapping sheet passes numeric codes through as-is
  (old 5 → new 5), even though the labels differ (Engaged vs Separated).
  See `MapMaritalStatus` in `Program.cs`.

---

## ~~PROFILES_PARTNERS_TB.IS_BANK_BORROWER — NOT NULL column with no source column~~ — resolved

- **Resolution**: stakeholder default `1` = لا/No. Code stores `IsBankBorrower
  = false` via `CodeBoolConverter` (false→1). See `MapPartner` in `Program.cs`.

---

## PROFILES_PARTNERS_TB.ID_NUM — rows with a partner name but no partner ID

- **Source**: `MF_CLIENT.PARTNER_NAME` present, `PARTNER_NATIONAL_ID` blank
- **Issue**: `ID_NUM` is `NOT NULL` on `PROFILES_PARTNERS_TB`, so these rows
  can't be inserted at all.
- **Current behavior**: the partner record is skipped entirely and logged to
  stderr as `[SKIP PARTNER] Row {n}: PARTNER_NAME='...' but
  PARTNER_NATIONAL_ID is missing`.
- **Needs**: confirm this skip behavior is acceptable, or decide on a
  fallback identifier.

---

## PROFILE_WORKS_TB.WORK_NATURE_ID — target code list unknown

- **Source**: `MF_CLIENT.PROFESSION_CODE`
- **Issue**: we have the *old* source codes for both companies, but not the
  *new/target* meaning of `WORK_NATURE_ID` itself:
  - **Asala** (employer type): 1=شركات/Companies, 2=أصحاب حرف ومهن حرة
    وتجار/Freelance/Own Business, 3=مؤسسات حكومية/Government institutions,
    4=أفراد قطاع عام/Public sector individuals, 5=أفراد قطاع خاص/Private
    sector individuals, 6=مؤسسات وجمعيات وأندية وهيئات اعتبارية/Institutions
    & societies, 7=تحويل/Transfer
  - **ACAD** (employment status): 1=Unemployed, 2=Freelance/Own Business,
    3=Employed, 4=Career Status
  - These are two different classification concepts (employer type vs.
    employment status), so this can't be a single shared lookup — it needs to
    be company-aware, similar to the `(COMPANY, COMPANY_BRANCH_CODE)` composite
    key used for `BranchId`.
- **Current behavior**: `WorkNatureId` is left `null` (never set in `MapWork`,
  `Program.cs`).
- **Needs**: the target `WORK_NATURE_ID` code list from the stakeholder, plus
  confirmation of how each company's codes map into it.

---

## ~~PROFILE_BANK_INFORMATION_TB — not populated at all yet~~ — resolved (placeholders)

- **Mapped**: `BANK_ACCOUNT_NUMBER → BANK_ACCOUNT_NO` when present.
- **Placeholders** (no source): `BANK_ID = 1`, `ACCOUNT_CURR_ID = 1`,
  `IBAN = MIG-{ProfileId}` (unique per profile for PK/unique index).
- **Nullable**: `BANK_ACCOUNT_NO = null` when the source account number is empty.
- See `MapBankInfo` in `Program.cs`. Replace placeholders when real bank/IBAN
  sources are available.

---
