"""
Target JSON schema for bill extraction.

Field selection mirrors exactly what the quote template (raw quote (2) sheet)
needs to be filled in — see excel_filler.py for the cell mapping. Meter Type
is deliberately excluded: that's filled in by Raj's team, not from the bill.
"""
from enum import Enum
from typing import List, Optional

from pydantic import BaseModel, Field


class ChargeUnit(str, Enum):
    KWH = "kWh"
    KVA = "kVA"
    KW = "kW"
    DAY = "day"
    MJ = "MJ"          # gas
    OTHER = "other"

    

class ChargeLine(BaseModel):
    """One row of the bill comparison table (Units / Description / Rate)."""

    description: str = Field(
        description=(
            "Charge name exactly as it would read on a quote, e.g. 'Daily charge', "
            "'Peak consumption', 'Off-peak consumption', 'Shoulder consumption', "
            "'Controlled load', 'Demand charge', 'Solar feed-in'."
        )
    )
    quantity: Optional[float] = Field(
        default=None,
        description="The billed quantity for this line (e.g. kWh used, or number of days for the daily charge).",
    )
    unit: Optional[ChargeUnit] = Field(
        default=None, description="Unit the quantity is measured in."
    )
    rate_before_discount: Optional[float] = Field(
        default=None,
        description=(
            "The GST-inclusive unit rate for this charge BEFORE any conditional/pay-on-time "
            "discount is applied, in dollars (e.g. 0.2557 for 25.57c/kWh, or 1.16314 for a "
            "$1.16314/day daily charge). Always convert cents to dollars."
        ),
    )
    conditional_discount_pct: Optional[float] = Field(
        default=None,
        description=(
            "If the bill shows a conditional / guaranteed / pay-on-time discount that applies "
            "to this charge, express it as a decimal fraction (e.g. 25% -> 0.25). Null if none."
        ),
    )
    is_credit: bool = Field(
        default=False,
        description="True if this line is a credit to the customer (e.g. solar feed-in), not a charge.",
    )

    is_usage: bool  # ADD THIS: True if it's consumption-based, False otherwise


class BillExtraction(BaseModel):
    customer_name: Optional[str] = Field(
        default=None, description="Account holder / customer name exactly as printed on the bill."
    )
    oc_number: Optional[str] = Field(
        default=None,
        description=(
            "Owners Corporation / Strata Plan number. Some bills (Strata/Owners Corporation "
            "customers) show this INSTEAD of a person's name in the customer-name position, "
            "typically followed by a mailing/postal address rather than the site address "
            "(e.g. 'OC 123456' or 'SP 12345' or 'Owners Corporation 123456'). When the "
            "customer-name field on the bill is actually an OC/SP number like this, put the "
            "number here (not in customer_name) and leave customer_name null. For ordinary "
            "bills with a normal person/business name, leave this null."
        ),
    )
    site_address: Optional[str] = Field(
        default=None, description="The supply address / site address (not the postal/billing address if different)."
    )
    nmi_or_mirn: Optional[str] = Field(
        default=None,
        description="National Metering Identifier (electricity, usually 10-11 digits) or MIRN (gas).",
    )
    distribution_region: Optional[str] = Field(
        default=None,
        description=(
            "The electricity/gas DISTRIBUTOR or network operator name (e.g. Powercor, AusNet Services, "
            "Ausgrid, Energex, SA Power Networks) — NOT the retailer. Only fill this if the distributor "
            "name is explicitly printed on the bill (e.g. under 'faults and emergencies'). Leave null if "
            "it isn't stated; do not guess from the address."
        ),
    )
    tariff_classification: Optional[str] = Field(
        default=None,
        description=(
            "The tariff type/classification as described on the bill, e.g. 'Anytime', "
            "'Time of Use' / 'Time of use + Controlled Load', . Only fill if the bill "
            "gives enough information to classify it — infer from the charge structure shown (e.g. if "
            "only one flat energy rate appears, it's 'Anytime'; if Off-Peak or sholulder rates appear, it's "
            "'Time of Use' if controlled load with Anytime or Time of use it will be + Controlled Load)."
        ),
    )
    current_energy_retailer: Optional[str] = Field(
        default=None, description="The retailer who issued this bill, e.g. 'Alinta Energy', '1st Energy'."
    )
    state: Optional[str] = Field(
        default=None, description="Australian state/territory abbreviation the site is in, e.g. VIC, NSW, QLD."
    )
    billing_period_days: Optional[int] = Field(
        default=None, description="Number of days covered by this bill's billing period."
    )
    charges: List[ChargeLine] = Field(
        default_factory=list,
        description="Every distinct charge/credit line item used to calculate the bill total.",
    )
    total_due: Optional[float] = Field(
        default=None, description="The total amount due on the bill, in dollars."
    )


def get_json_schema() -> dict:
    """JSON schema to hand to Gemini's response_json_schema config."""
    return BillExtraction.model_json_schema()


class ConsolidatedBillExtraction(BaseModel):
    """
    Target schema for a single PDF that contains bills for MULTIPLE sites/meters
    (e.g. a Strata/Owners Corporation portfolio bill, or several bills merged into
    one PDF for a single consolidated quote).
    """

    bills: List[BillExtraction] = Field(
        default_factory=list,
        description=(
            "One entry per distinct site/meter/NMI found in the PDF, in the order they "
            "appear. Treat each site as a fully separate bill for extraction purposes — "
            "do not merge charges from different NMIs/sites into one entry, even if they "
            "share the same customer_name/oc_number or appear on the same statement page."
        ),
    )


def get_consolidated_json_schema() -> dict:
    """JSON schema to hand to Gemini's response_json_schema config for multi-site PDFs."""
    return ConsolidatedBillExtraction.model_json_schema()