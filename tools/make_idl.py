"""Replaces the @IID:name@ and @CLSID:name@ placeholders of tlb2idl's output by
WFC's own identifiers (name-based UUIDs, so the same text always gives the same
ids) and writes the IDL.

usage: tlb2idl msado15.dll | python make_idl.py > ../src/adodb.idl
"""
import re
import sys
import uuid

NAMESPACE = uuid.UUID("6ba7b811-9dad-11d1-80b4-00c04fd430c8")  # the URL namespace
BASE = "https://github.com/Thewafflication/wfc-adodb/"


def make(kind, name):
    if kind == "CLSID":
        text = BASE + name  # the ids the registration already uses
    else:
        text = BASE + kind.lower() + "/" + name
    return str(uuid.uuid5(NAMESPACE, text)).upper()


text = sys.stdin.read()
text = re.sub(r"@(IID|CLSID):([A-Za-z_0-9]+)@",
              lambda m: make(m.group(1), m.group(2)), text)
sys.stdout.write(text)
