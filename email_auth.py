import random
import resend

def generate_otp() -> str:
    return str(random.randint(100000, 999999))

def send_otp_email(resend_api_key: str, receiver_email: str, otp_code: str) -> tuple[bool, str]:
    try:
        resend.api_key = resend_api_key
        resend.Emails.send({
            "from": "BusinessGem <onboarding@resend.dev>",
            "to": receiver_email,
            "subject": "[비즈니스 분석 Gem] 로그인 인증 패스코드",
            "html": f"""
            <h3>안녕하세요,</h3>
            <p>비즈니스 분석 Gem 로그인을 위한 인증 패스코드입니다.</p>
            <h2>[ 인증 패스코드 : {otp_code} ]</h2>
            <p>- 해당 패스코드는 5분간 유효합니다.</p>
            """
        })
        return True, "성공"
    except Exception as e:
        return False, str(e)
