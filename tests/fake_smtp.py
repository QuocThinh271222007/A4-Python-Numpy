"""SMTP giả lập để thử nghiệm: in thư nhận được; từ chối bad@ (550); báo lỗi tạm 451 lần đầu với giang@."""
import asyncio, sys
from aiosmtpd.controller import Controller
seen = {"giang": 0}
class H:
    async def handle_RCPT(self, server, session, envelope, address, rcpt_options):
        if address.startswith("phong"):
            return "550 5.1.1 mailbox unavailable"
        if address.startswith("giang"):
            seen["giang"] += 1
            if seen["giang"] == 1:
                return "451 4.3.0 try again later"
        envelope.rcpt_tos.append(address); return "250 OK"
    async def handle_DATA(self, server, session, envelope):
        subj = [l for l in envelope.content.decode("utf8","replace").splitlines() if l.lower().startswith("subject")]
        print("RECV to=%s %s size=%d" % (envelope.rcpt_tos, subj[:1], len(envelope.content)), flush=True)
        return "250 queued"
c = Controller(H(), hostname="127.0.0.1", port=2525); c.start()
print("ready", flush=True)
try: asyncio.get_event_loop().run_forever()
except KeyboardInterrupt: pass
