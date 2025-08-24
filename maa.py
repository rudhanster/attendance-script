def matches_event_text(txt: str, code: str, sem: str, sec: str, sess_ignored: str | None) -> bool:
    """
    - Section can be 'B' or 'B-1'.
    - If section is a single letter (e.g., 'B'), DO NOT match 'B-1', 'B-2', etc.
      Only accept 'Sec B', 'Section B', etc.
    - Session is ignored by design.
    """
    T = _norm(txt).upper()
    reasons = []   # collect why it fails

    ok = True

    if code:
        if code.upper() not in T:
            reasons.append(f"Course code '{code}' not in '{T}'")
            ok = False

    if sem:
        if f"SEMESTER {sem.upper()}" not in T:
            reasons.append(f"Semester '{sem}' not in '{T}'")
            ok = False

    if sec:
        secU = sec.upper().strip()
        if "-" in secU:
            # Exact like B-1 → require exact match as a whole word
            if not re.search(rf"\b{re.escape(secU)}\b", T):
                reasons.append(f"Section '{secU}' not in '{T}'")
                ok = False
        else:
            # Only match explicit section tokens: SEC B / SECTION B
            patt = rf"\bSEC(?:TION)?\.?\s*[:\-]?\s*{re.escape(secU)}\b(?!-)"
            if not re.search(patt, T):
                reasons.append(f"Section '{secU}' not matched explicitly in '{T}'")
                ok = False

    if not ok:
        print(f"❌ Event text mismatch: {txt}  →  {', '.join(reasons)}")

    return ok

