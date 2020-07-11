"""Sweep through various suppression options to find best true vs. false positive tradeoff."""

import os
import re
import subprocess

maxCategories = "3"
reportingThreshold = 0

opts = [
#    "suppressDifferentReferentCount",
    "suppressFatFix",
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

# HACK
lenopts = 0

for reportingThreshold in [0.5 * i for i in range(45,55)]:
    results = {}
    for i in range(2**lenopts):
        thestr = f":0{lenopts}b"
        bitstring = ("{" + thestr + "}").format(i)
        # print(bitstring)
        for b in range(lenopts):
            optionvector[b] = (bitstring[b] == "1")
        # print(optionvector)
        optionstr = ""
        for b in range(lenopts):
            if optionvector[b]:
                optionstr += " " + "--" + opts[b] + "=1"

        optionstr = "--suppressFatFix=1 --suppressRecurrentFormula=1 --suppressOneExtraConstant=1 --suppressNumberOfConstantsMismatch=1 --suppressBothConstants=1 --suppressOneIsAllConstants=1"
        cmdline = "node ../dist/excelint-cli.js --maxCategories " + maxCategories + " --reportingThreshold " + str(reportingThreshold) + " --directory subjects_xlsx " + optionstr
        # print(cmdline)
        # cmdline = "ls -l"
        r = subprocess.run(cmdline, shell=True, capture_output=True)
        truePositives = 0
        falsePositives = 0
        for line in r.stdout.decode().split('\n'):
            # Try to match out true positives and false positives
            if (line.find("ExceLint true positives") >= 0):
                truePositives = int(re.findall("[0-9]+", line)[0])
            if (line.find("ExceLint false positives") >= 0):
                falsePositives = int(re.findall("[0-9]+", line)[0])
        if truePositives + falsePositives == 0:
            continue
        effectiveness = float(truePositives) / (float(truePositives) + float(falsePositives))
        print(effectiveness, reportingThreshold, optionstr, truePositives, falsePositives, flush=True)
        results[optionstr] = [truePositives, falsePositives]

    print(results, flush=True)

#cmdline = "ls -l"
#r = subprocess.run(cmdline, shell=True, capture_output=True)
#for lines in r.stdout.decode().split('\n'):
#    print(lines)



