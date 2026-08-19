# Sector guide: regional healthcare

How a health service (especially a regional one with lean corporate
capacity) gets the most from this template library, and the boundaries
that keep a SharePoint deployment on the right side of clinical and
statutory systems.

## The two boundaries (read these first)

**1. Clinical data does not belong in SharePoint lists.** Nothing
patient-identifiable (no clinical records, no patient-level incident
detail, no health information) goes into any template here. The templates
manage *corporate and clinical-governance processes*; patient data lives in
your clinical systems under their controls. Where a register touches a
clinical event (a complaint about care, an improvement from an incident),
record the substance at the process level and reference the clinical
system's identifier, never the clinical content.

The same boundary applies to **research data**, and it is worth naming
separately because the erosion there is the reasonable-sounding column.
A research ethics register holds project-level metadata only: what was
approved, by whom, until when, and what is owed. Not participant names, not
identifiers, and **not recruitment or consent counts, not even totals**: an
aggregate count is not identifiable, and it is also the column before "which
twelve". Those numbers belong in the study file or the clinical trial
management system, which have the controls for them.

**2. Statutory and mandated systems win.** Victorian health services (and
equivalents elsewhere) run mandated systems for clinical incidents
(VHIMS/RiskMan-class), and many run enterprise risk in the same platform.
A SharePoint list never replaces a mandated system:

- **incident-management** here = *corporate/non-clinical* incidents
  (security, IT, facilities, process failures). Clinical incidents go to
  the statutory system, and nowhere else.
- **complaints-feedback** here works well as the consumer-feedback
  register for services not mandated onto a specific platform, but check
  your health-complaints scheme's requirements first, and record
  referrals to the commissioner/ombudsman as the template's governance
  describes.
- **risk-register** here suits committee/departmental/project risk tiers;
  if your enterprise register is mandated elsewhere, use this for the
  local tiers and escalate material risks upward per your framework
  (the template's RiskMan-style escalation hook pattern exists for
  exactly that).
- **opportunities-register** is only a de-identified project-to-business
  routing layer. Its stop gate sends current harm, incidents, complaints,
  privacy/cyber matters and statutory concerns to their mandated systems;
  only a safe process observation and destination reference remain here.

## Template map for a health service

| Health-service need | Template | Notes |
| --- | --- | --- |
| Practitioner credentialing & scope (NSQHS Std 1) | credentialing-register | The register accreditors ask for first |
| Biomed / test-and-tag / fire maintenance | equipment-maintenance | Evidence-linked service history |
| Fridge temps, trolley checks, cleaning rounds | routine-checks | Cold-chain evidence; kills the paper sheets |
| Switchboard: code log, message book, key register | switchboard-log | The emergency-planning committee's evidence base |
| NSQHS / aged-care standards evidence | compliance-obligations | One slice per standard, end-to-end |
| Volunteers with police/WWCC discipline | volunteer-register | The expiry sweep is the control |
| Grants and their acquittals | grants-register | Regional services live on these |
| Consumer feedback (Std 2 partnering) | complaints-feedback | Statutory-scheme check first |
| Quality improvement (Std 1 CQI) | improvement-register | Feed it from feedback, audits, incidents |
| Out-of-scope problems discovered by projects | opportunities-register | One-minute capture; route to existing controls before assessing anything |
| Clinical audit actions / accreditation findings | audit-actions | Recommendations to closure with evidence |
| Research & QI projects sent to a partner HREC | research-ethics-register-simple | The single-list register for a service referring to a partner's HREC; a service with its own research office wants the multi-list shape instead. Two separate gates (ethics approval and site authorisation) on one row |
| Where digitised records may be filed | records-digitisation | One row per platform: can it keep a record, or only store a file. The register a digitisation program needs before it scans anything |
| Committee meetings, decisions, actions | meeting-actions | Governance-lite for every committee |
| Corporate risk tiers | risk-register | See boundary 2 |
| Everything else | the general library | Assets, contracts, onboarding, training, service requests... |

## A pragmatic first-90-days sequence

1. **Week 1, quick wins that build trust**: meeting-actions for two
   committees; routine-checks for the vaccine fridges.
2. **Month 1, the compliance spine**: credentialing-register load;
   equipment-maintenance schedule load; volunteer-register load.
3. **Month 2-3, the improvement engine**: process-register inventory
   workshops; compliance-obligations one-standard slice;
   improvement-register opened and fed.
4. **Then**: let the process-register worklist drive what gets digitised
   next. That's the point of it.

## Privacy defaults for health deployments

The restrictive templates (complaints-feedback, volunteer-register and
opportunities-register) ship with no general-staff access; keep it that way.
For everything else,
remember site membership *is* the audience. A "whole of org" site makes
every Read grant organisation-wide. Registers holding staff professional
data (credentialing) are read-wide by operational design but deserve a
deliberate membership decision, recorded in their governance doc.
