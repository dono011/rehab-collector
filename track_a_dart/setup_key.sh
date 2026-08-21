#!/bin/bash
# DART 인증키를 서버에 넣는다. (2026-08-21)
#
# 사장님이 직접 실행하신다. 키는 건우(클로드)를 거치지 않는다.
# 화면에 키를 찍지 않고, 명령 기록에도 남지 않게 물어보는 방식으로 받는다.
#
#   ssh -t -i C:\Users\1\.ssh\id_ed25519 root@187.127.207.110 /root/rehab-collector/track_a_dart/setup_key.sh
#
# 키를 인자로 넘겨도 되지만(자동화용) 그러면 기록에 남는다.

set -u
ENV_FILE="/root/rehab-collector/.env"
TRACK_A="/root/rehab-collector/track_a_dart"

KEY="${1:-}"
if [ -z "$KEY" ]; then
    if [ -t 0 ]; then
        echo
        echo "DART 인증키(영문+숫자 40자리)를 붙여넣고 엔터를 누르십시오."
        echo "붙여넣어도 화면에는 안 보입니다. 그냥 엔터 치시면 됩니다."
        echo
        printf "인증키: "
        read -r -s KEY
        echo
    else
        # 클립보드에서 흘려보낸 경우 (바탕화면 DART_key.bat)
        echo "클립보드에서 인증키를 읽었습니다."
        read -r KEY
    fi
fi

KEY="$(echo "$KEY" | tr -d '[:space:]')"

if ! echo "$KEY" | grep -qE '^[0-9a-fA-F]{40}$'; then
    echo "[실패] 키 모양이 맞지 않습니다."
    echo "        받은 길이: ${#KEY}자  (40자리여야 합니다)"
    echo "        opendart.fss.or.kr → 인증키 신청/관리 → 오픈API 이용현황 에서"
    echo "        40자리를 통째로 복사해 다시 시도하십시오."
    exit 1
fi

umask 077
touch "$ENV_FILE"
chmod 600 "$ENV_FILE"
grep -v '^DART_API_KEY=' "$ENV_FILE" > "$ENV_FILE.tmp" 2>/dev/null || true
mv "$ENV_FILE.tmp" "$ENV_FILE"
echo "DART_API_KEY=$KEY" >> "$ENV_FILE"
chmod 600 "$ENV_FILE"

echo "[저장] $ENV_FILE (이 파일은 root만 읽을 수 있습니다)"
echo
echo "[확인] 금융감독원에 연결해 봅니다..."
cd "$TRACK_A" || exit 1
if DART_API_KEY="$KEY" timeout 60 /usr/bin/python3 collect_rehab.py --test; then
    echo
    echo "=============================================="
    echo " 성공했습니다. 건우에게 '됐다'고 알려주십시오."
    echo "=============================================="
else
    echo
    echo "[실패] 키는 저장했으나 연결이 안 됩니다."
    echo "        승인이 아직 안 끝났을 수 있습니다. 위 메시지를 건우에게 보여주십시오."
    exit 1
fi
