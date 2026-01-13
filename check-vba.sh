#!/bin/bash
# check-vba.sh - Static analysis for VBA modules
# Usage: ./check-vba.sh [--quick]

QUICK=false
[[ "$1" == "--quick" ]] && QUICK=true

RED='\033[0;31m'
YEL='\033[0;33m'
GRN='\033[0;32m'
NC='\033[0m'

ERRORS=0
WARNS=0

err() { echo -e "${RED}ERROR:${NC} $1"; ((ERRORS++)); }
warn() { echo -e "${YEL}WARN:${NC} $1"; ((WARNS++)); }

echo "=== VBA Static Check ==="
echo ""

# Find all .bas and .cls files
FILES=$(find . -maxdepth 1 \( -name "*.bas" -o -name "*.cls" \) 2>/dev/null)

if [[ -z "$FILES" ]]; then
    echo "No .bas or .cls files found"
    exit 0
fi

# 1. Build artifacts (cause compile errors)
echo "Checking for build artifacts..."
for f in $FILES; do
    if grep -q "^Attribute VB_" "$f"; then
        err "$f contains Attribute VB_* lines (remove before import)"
    fi
    if grep -q "^VERSION 1.0 CLASS" "$f"; then
        err "$f contains VERSION header (remove before import)"
    fi
done

# 2. Option Explicit check
echo "Checking Option Explicit..."
for f in $FILES; do
    if ! grep -q "^Option Explicit" "$f"; then
        err "$f missing Option Explicit"
    fi
done

# 3. Sub/End Sub matching (handle single-line subs)
echo "Checking Sub/End Sub balance..."
for f in $FILES; do
    # Count multi-line subs (Sub on line without End Sub)
    MULTI_SUBS=$(grep -E "^(Public |Private )?Sub " "$f" | grep -cv "End Sub")
    MULTI_SUBS=${MULTI_SUBS:-0}
    # Count End Sub on its own line
    END_SUBS=$(grep -c "^End Sub$" "$f")
    END_SUBS=${END_SUBS:-0}
    if [[ "$MULTI_SUBS" != "$END_SUBS" ]]; then
        err "$f Sub/End Sub mismatch ($MULTI_SUBS multi-line subs, $END_SUBS End Subs)"
    fi
done

# 4. Function/End Function matching
echo "Checking Function/End Function balance..."
for f in $FILES; do
    FUNCS=$(grep -cE "^(Public |Private )?Function " "$f")
    FUNCS=${FUNCS:-0}
    ENDS=$(grep -c "^End Function$" "$f")
    ENDS=${ENDS:-0}
    if [[ "$FUNCS" != "$ENDS" ]]; then
        err "$f Function/End Function mismatch ($FUNCS funcs, $ENDS ends)"
    fi
done

# Quick mode stops here
if $QUICK; then
    echo ""
    if [[ $ERRORS -gt 0 ]]; then
        echo -e "${RED}FAILED:${NC} $ERRORS error(s)"
        exit 1
    else
        echo -e "${GRN}PASSED${NC} (quick mode)"
        exit 0
    fi
fi

# 5. Public subs without error handling (warn only)
echo "Checking error handling in public subs..."
for f in $FILES; do
    BASE=$(basename "$f")
    [[ "$BASE" == "Tests.bas" || "$BASE" == "Setup.bas" || "$BASE" == "Scenarios.bas" || "$BASE" == "Validate.bas" ]] && continue

    while IFS=: read -r line content; do
        [[ -z "$line" ]] && continue
        HAS_ERROR=$(sed -n "${line},$((line+10))p" "$f" | grep -c "On Error")
        if [[ "${HAS_ERROR:-0}" == "0" ]]; then
            SUB_NAME=$(echo "$content" | sed 's/Public Sub \([^(]*\).*/\1/')
            warn "$f:$line $SUB_NAME lacks error handling"
        fi
    done < <(grep -n "^Public Sub " "$f" 2>/dev/null)
done

# 6. Schema constant consistency
echo "Checking Schema constant usage..."
if [[ -f "Schema.bas" ]]; then
    for f in $FILES; do
        [[ "$f" == "./Schema.bas" ]] && continue
        while read -r const; do
            [[ -z "$const" ]] && continue
            CONST_NAME="${const#Schema.}"
            if ! grep -q "^Public Const $CONST_NAME " Schema.bas 2>/dev/null; then
                err "$f references undefined $const"
            fi
        done < <(grep -oE "Schema\.[A-Z][A-Z0-9_]+" "$f" 2>/dev/null | sort -u)
    done
fi

# 7. Duplicate function names across modules
echo "Checking for duplicate function names..."
ALL_FUNCS=$(grep -hE "^(Public |Private )?(Sub|Function) " $FILES 2>/dev/null | \
    sed 's/.*\(Sub\|Function\) \([^(:]*\).*/\2/' | sort)
DUPES=$(echo "$ALL_FUNCS" | uniq -d)
for d in $DUPES; do
    [[ -n "$d" ]] && warn "Duplicate name: $d"
done

# Summary
echo ""
echo "=== Summary ==="
if [[ $ERRORS -gt 0 ]]; then
    echo -e "${RED}FAILED:${NC} $ERRORS error(s), $WARNS warning(s)"
    exit 1
elif [[ $WARNS -gt 0 ]]; then
    echo -e "${YEL}PASSED with warnings:${NC} $WARNS warning(s)"
    exit 0
else
    echo -e "${GRN}PASSED:${NC} All checks clean"
    exit 0
fi
