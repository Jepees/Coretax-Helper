"""
XML generator for Coretax BP21, BPMP, BPA1, and BPPU formats.
Generates XML strings matching the official Coretax import schema.
"""

import xml.etree.ElementTree as ET
from datetime import date, datetime


XSI_NAMESPACE = "http://www.w3.org/2001/XMLSchema-instance"


def _format_date(value) -> str:
    """Convert a date/datetime value to YYYY-MM-DD string."""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    elif isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    elif isinstance(value, str):
        # Already a string, try to normalize
        value = value.strip()
        # Handle common formats
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        return value
    return str(value)


def _format_number(value) -> str:
    """Format a number: remove decimals if whole number, round to avoid float artifacts."""
    if value is None:
        return "0"
    num = round(float(value), 2)
    if num == int(num):
        return str(int(num))
    return str(num)


def _indent_xml(elem: ET.Element, level: int = 0) -> None:
    """Add pretty-print indentation to XML tree (using tabs like Coretax)."""
    indent = "\n" + "\t" * level
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = indent + "\t"
        if not elem.tail or not elem.tail.strip():
            elem.tail = indent
        for child in elem:
            _indent_xml(child, level + 1)
        if not child.tail or not child.tail.strip():
            child.tail = indent
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = indent


def generate_bp21_xml(tin: str, rows: list[dict]) -> str:
    """
    Generate BP21 XML string.

    Args:
        tin: NPWP Pemotong (16 digits)
        rows: List of dicts with keys matching BP21 XML tags

    Returns:
        XML string ready for download
    """
    root = ET.Element("Bp21Bulk")
    root.set("xmlns:xsi", XSI_NAMESPACE)

    tin_elem = ET.SubElement(root, "TIN")
    tin_elem.text = str(tin)

    list_elem = ET.SubElement(root, "ListOfBp21")

    for row in rows:
        bp21 = ET.SubElement(list_elem, "Bp21")

        # Fields in order matching the official Coretax schema
        fields = [
            ("TaxPeriodMonth", _format_number(row.get("TaxPeriodMonth", ""))),
            ("TaxPeriodYear", _format_number(row.get("TaxPeriodYear", ""))),
            ("CounterpartTin", str(row.get("CounterpartTin", ""))),
            ("IDPlaceOfBusinessActivityOfIncomeRecipient", str(row.get("IDPlaceOfBusinessActivityOfIncomeRecipient", ""))),
            ("StatusTaxExemption", str(row.get("StatusTaxExemption", ""))),
            ("TaxCertificate", str(row.get("TaxCertificate", ""))),
            ("TaxObjectCode", str(row.get("TaxObjectCode", ""))),
            ("Gross", _format_number(row.get("Gross", 0))),
            ("Deemed", _format_number(row.get("Deemed", 100))),
            ("Rate", _format_number(row.get("Rate", 0))),
            ("Document", str(row.get("Document", ""))),
            ("DocumentNumber", str(row.get("DocumentNumber", ""))),
            ("DocumentDate", _format_date(row.get("DocumentDate", ""))),
            ("IDPlaceOfBusinessActivity", str(row.get("IDPlaceOfBusinessActivity", ""))),
            ("WithholdingDate", _format_date(row.get("WithholdingDate", ""))),
        ]

        for tag, value in fields:
            elem = ET.SubElement(bp21, tag)
            elem.text = value

    _indent_xml(root)

    xml_declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    xml_body = ET.tostring(root, encoding="unicode")

    return xml_declaration + xml_body


def generate_bpmp_xml(tin: str, rows: list[dict]) -> str:
    """
    Generate BPMP XML string.

    Args:
        tin: NPWP Pemotong (16 digits)
        rows: List of dicts with keys matching BPMP XML tags

    Returns:
        XML string ready for download
    """
    root = ET.Element("MmPayrollBulk")
    root.set("xmlns:xsi", XSI_NAMESPACE)

    tin_elem = ET.SubElement(root, "TIN")
    tin_elem.text = str(tin)

    list_elem = ET.SubElement(root, "ListOfMmPayroll")

    for row in rows:
        payroll = ET.SubElement(list_elem, "MmPayroll")

        # TaxPeriodMonth
        elem = ET.SubElement(payroll, "TaxPeriodMonth")
        elem.text = _format_number(row.get("TaxPeriodMonth", ""))

        # TaxPeriodYear
        elem = ET.SubElement(payroll, "TaxPeriodYear")
        elem.text = _format_number(row.get("TaxPeriodYear", ""))

        # CounterpartOpt
        elem = ET.SubElement(payroll, "CounterpartOpt")
        elem.text = str(row.get("CounterpartOpt", "Resident"))

        # CounterpartPassport — xsi:nil="true" if Resident (no passport)
        counterpart_opt = str(row.get("CounterpartOpt", "Resident"))
        passport = row.get("CounterpartPassport")
        elem = ET.SubElement(payroll, "CounterpartPassport")
        if counterpart_opt == "Resident" or not passport:
            elem.set("xsi:nil", "true")
        else:
            elem.text = str(passport)

        # CounterpartTin
        elem = ET.SubElement(payroll, "CounterpartTin")
        elem.text = str(row.get("CounterpartTin", ""))

        # StatusTaxExemption
        elem = ET.SubElement(payroll, "StatusTaxExemption")
        elem.text = str(row.get("StatusTaxExemption", ""))

        # Position
        elem = ET.SubElement(payroll, "Position")
        elem.text = str(row.get("Position", ""))

        # TaxCertificate
        elem = ET.SubElement(payroll, "TaxCertificate")
        elem.text = str(row.get("TaxCertificate", ""))

        # TaxObjectCode
        elem = ET.SubElement(payroll, "TaxObjectCode")
        elem.text = str(row.get("TaxObjectCode", ""))

        # Gross
        elem = ET.SubElement(payroll, "Gross")
        elem.text = _format_number(row.get("Gross", 0))

        # Rate
        elem = ET.SubElement(payroll, "Rate")
        elem.text = _format_number(row.get("Rate", 0))

        # IDPlaceOfBusinessActivity
        elem = ET.SubElement(payroll, "IDPlaceOfBusinessActivity")
        elem.text = str(row.get("IDPlaceOfBusinessActivity", ""))

        # WithholdingDate
        elem = ET.SubElement(payroll, "WithholdingDate")
        elem.text = _format_date(row.get("WithholdingDate", ""))

    _indent_xml(root)

    xml_declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    xml_body = ET.tostring(root, encoding="unicode")

    return xml_declaration + xml_body


def generate_bpa1_xml(tin: str, rows: list[dict]) -> str:
    """
    Generate BPA1 XML string (Bukti Potong Tahunan A1).

    Args:
        tin: NPWP Pemotong (16 digits)
        rows: List of dicts with keys matching BPA1 XML tags

    Returns:
        XML string ready for download
    """
    root = ET.Element("A1Bulk")
    root.set("xmlns:xsi", XSI_NAMESPACE)

    tin_elem = ET.SubElement(root, "TIN")
    tin_elem.text = str(tin)

    list_elem = ET.SubElement(root, "ListOfA1")

    for row in rows:
        a1 = ET.SubElement(list_elem, "A1")

        # WorkForSecondEmployer
        elem = ET.SubElement(a1, "WorkForSecondEmployer")
        elem.text = str(row.get("WorkForSecondEmployer", "No"))

        # TaxPeriodMonthStart
        elem = ET.SubElement(a1, "TaxPeriodMonthStart")
        elem.text = _format_number(row.get("TaxPeriodMonthStart", ""))

        # TaxPeriodMonthEnd
        elem = ET.SubElement(a1, "TaxPeriodMonthEnd")
        elem.text = _format_number(row.get("TaxPeriodMonthEnd", ""))

        # TaxPeriodYear
        elem = ET.SubElement(a1, "TaxPeriodYear")
        elem.text = _format_number(row.get("TaxPeriodYear", ""))

        # CounterpartOpt
        elem = ET.SubElement(a1, "CounterpartOpt")
        elem.text = str(row.get("CounterpartOpt", "Resident"))

        # CounterpartPassport — xsi:nil if Resident or empty
        passport = row.get("CounterpartPassport")
        counterpart_opt = str(row.get("CounterpartOpt", "Resident"))
        elem = ET.SubElement(a1, "CounterpartPassport")
        if counterpart_opt == "Resident" or not passport:
            elem.set("xsi:nil", "true")
        else:
            elem.text = str(passport)

        # CounterpartTin
        elem = ET.SubElement(a1, "CounterpartTin")
        elem.text = str(row.get("CounterpartTin", ""))

        # TaxExemptOpt
        elem = ET.SubElement(a1, "TaxExemptOpt")
        elem.text = str(row.get("TaxExemptOpt", ""))

        # StatusOfWithholding
        elem = ET.SubElement(a1, "StatusOfWithholding")
        elem.text = str(row.get("StatusOfWithholding", ""))

        # CounterpartPosition
        elem = ET.SubElement(a1, "CounterpartPosition")
        elem.text = str(row.get("CounterpartPosition", ""))

        # TaxObjectCode
        elem = ET.SubElement(a1, "TaxObjectCode")
        elem.text = str(row.get("TaxObjectCode", ""))

        # NumberOfMonths
        elem = ET.SubElement(a1, "NumberOfMonths")
        elem.text = _format_number(row.get("NumberOfMonths", 0))

        # SalaryPensionJhtTht
        elem = ET.SubElement(a1, "SalaryPensionJhtTht")
        elem.text = _format_number(row.get("SalaryPensionJhtTht", 0))

        # GrossUpOpt
        elem = ET.SubElement(a1, "GrossUpOpt")
        elem.text = str(row.get("GrossUpOpt", "No"))

        # IncomeTaxBenefit
        elem = ET.SubElement(a1, "IncomeTaxBenefit")
        elem.text = _format_number(row.get("IncomeTaxBenefit", 0))

        # OtherBenefit
        elem = ET.SubElement(a1, "OtherBenefit")
        elem.text = _format_number(row.get("OtherBenefit", 0))

        # Honorarium
        elem = ET.SubElement(a1, "Honorarium")
        elem.text = _format_number(row.get("Honorarium", 0))

        # InsurancePaidByEmployer
        elem = ET.SubElement(a1, "InsurancePaidByEmployer")
        elem.text = _format_number(row.get("InsurancePaidByEmployer", 0))

        # Natura
        elem = ET.SubElement(a1, "Natura")
        elem.text = _format_number(row.get("Natura", 0))

        # TantiemBonusThr
        elem = ET.SubElement(a1, "TantiemBonusThr")
        elem.text = _format_number(row.get("TantiemBonusThr", 0))

        # PensionContributionJhtThtFee
        elem = ET.SubElement(a1, "PensionContributionJhtThtFee")
        elem.text = _format_number(row.get("PensionContributionJhtThtFee", 0))

        # Zakat
        elem = ET.SubElement(a1, "Zakat")
        elem.text = _format_number(row.get("Zakat", 0))

        # PrevWhTaxSlip — xsi:nil if empty
        prev_wh = row.get("PrevWhTaxSlip")
        elem = ET.SubElement(a1, "PrevWhTaxSlip")
        if not prev_wh:
            elem.set("xsi:nil", "true")
        else:
            elem.text = str(prev_wh)

        # TaxCertificate
        elem = ET.SubElement(a1, "TaxCertificate")
        elem.text = str(row.get("TaxCertificate", "N/A"))

        # Article21IncomeTax
        elem = ET.SubElement(a1, "Article21IncomeTax")
        elem.text = _format_number(row.get("Article21IncomeTax", 0))

        # IDPlaceOfBusinessActivity
        elem = ET.SubElement(a1, "IDPlaceOfBusinessActivity")
        elem.text = str(row.get("IDPlaceOfBusinessActivity", ""))

        # WithholdingDate
        elem = ET.SubElement(a1, "WithholdingDate")
        elem.text = _format_date(row.get("WithholdingDate", ""))

    _indent_xml(root)

    xml_declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    xml_body = ET.tostring(root, encoding="unicode")

    return xml_declaration + xml_body


def generate_bppu_xml(tin: str, rows: list[dict]) -> str:
    """
    Generate BPPU XML string (Bukti Potong PPh Unifikasi).

    Args:
        tin: NPWP Pemotong (16 digits)
        rows: List of dicts with keys matching BPPU XML tags

    Returns:
        XML string ready for download
    """
    root = ET.Element("BpuBulk")
    root.set("xmlns:xsi", XSI_NAMESPACE)

    tin_elem = ET.SubElement(root, "TIN")
    tin_elem.text = str(tin)

    list_elem = ET.SubElement(root, "ListOfBpu")

    for row in rows:
        bpu = ET.SubElement(list_elem, "Bpu")

        # Fields in order matching the official Coretax schema
        # TaxPeriodMonth
        elem = ET.SubElement(bpu, "TaxPeriodMonth")
        elem.text = _format_number(row.get("TaxPeriodMonth", ""))

        # TaxPeriodYear
        elem = ET.SubElement(bpu, "TaxPeriodYear")
        elem.text = _format_number(row.get("TaxPeriodYear", ""))

        # CounterpartTin
        elem = ET.SubElement(bpu, "CounterpartTin")
        elem.text = str(row.get("CounterpartTin", ""))

        # IDPlaceOfBusinessActivityOfIncomeRecipient
        elem = ET.SubElement(bpu, "IDPlaceOfBusinessActivityOfIncomeRecipient")
        elem.text = str(row.get("IDPlaceOfBusinessActivityOfIncomeRecipient", ""))

        # TaxCertificate
        elem = ET.SubElement(bpu, "TaxCertificate")
        elem.text = str(row.get("TaxCertificate", "N/A"))

        # TaxObjectCode
        elem = ET.SubElement(bpu, "TaxObjectCode")
        elem.text = str(row.get("TaxObjectCode", ""))

        # TaxBase
        elem = ET.SubElement(bpu, "TaxBase")
        elem.text = _format_number(row.get("TaxBase", 0))

        # Rate
        elem = ET.SubElement(bpu, "Rate")
        elem.text = _format_number(row.get("Rate", 0))

        # Document
        elem = ET.SubElement(bpu, "Document")
        elem.text = str(row.get("Document", ""))

        # DocumentNumber
        elem = ET.SubElement(bpu, "DocumentNumber")
        elem.text = str(row.get("DocumentNumber", ""))

        # DocumentDate
        elem = ET.SubElement(bpu, "DocumentDate")
        elem.text = _format_date(row.get("DocumentDate", ""))

        # IDPlaceOfBusinessActivity
        elem = ET.SubElement(bpu, "IDPlaceOfBusinessActivity")
        elem.text = str(row.get("IDPlaceOfBusinessActivity", ""))

        # GovTreasurerOpt
        elem = ET.SubElement(bpu, "GovTreasurerOpt")
        elem.text = str(row.get("GovTreasurerOpt", "N/A"))

        # SP2DNumber — xsi:nil if empty
        sp2d = row.get("SP2DNumber")
        elem = ET.SubElement(bpu, "SP2DNumber")
        if not sp2d:
            elem.set("xsi:nil", "true")
        else:
            elem.text = str(sp2d)

        # WithholdingDate
        elem = ET.SubElement(bpu, "WithholdingDate")
        elem.text = _format_date(row.get("WithholdingDate", ""))

    _indent_xml(root)

    xml_declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    xml_body = ET.tostring(root, encoding="unicode")

    return xml_declaration + xml_body
