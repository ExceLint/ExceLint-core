"""Sweep through various suppression options to find best true vs. false positive tradeoff."""

import os
import re
import subprocess

maxCategories = "3"
reportingThreshold = "2"

opts = [
#    "suppressDifferentReferentCount",
    "suppressRecurrentFormula",
    "suppressOneExtraConstant",
    "suppressNumberOfConstantsMismatch",
    "suppressBothConstants",
    "suppressOneIsAllConstants",
#    "suppressR1C1Mismatch",
#    "suppressAbsoluteRefMismatch",
#    "suppressOffAxisReference"
]

# opts = ["suppressDifferentReferentCount"]

lenopts = len(opts)

optionvector = [0] * lenopts

results = {}

for i in range(2**lenopts):
    str = f":0{lenopts}b"
    bitstring = ("{" + str + "}").format(i)
    # print(bitstring)
    for b in range(lenopts):
        optionvector[b] = (bitstring[b] == "1")
    # print(optionvector)
    optionstr = ""
    for b in range(lenopts):
        if optionvector[b]:
            optionstr += " " + "--" + opts[b] + "=1"
        
    cmdline = "node ../dist/excelint-cli.js --maxCategories " + maxCategories + " --reportingThreshold " + reportingThreshold + " --directory subjects_xlsx " + optionstr
    # print(cmdline)
    # cmdline = "ls -l"
    r = subprocess.run(cmdline, shell=True, capture_output=True)
    for line in r.stdout.decode().split('\n'):
        # Try to match out true positives and false positives
        if (line.find("ExceLint true positives") >= 0):
            truePositives = int(re.findall("[0-9]+", line)[0])
        if (line.find("ExceLint false positives") >= 0):
            falsePositives = int(re.findall("[0-9]+", line)[0])
    print(optionstr, truePositives, falsePositives, flush=True)
    results[optionstr] = [truePositives, falsePositives]

print(results, flush=True)

#cmdline = "ls -l"
#r = subprocess.run(cmdline, shell=True, capture_output=True)
#for lines in r.stdout.decode().split('\n'):
#    print(lines)



