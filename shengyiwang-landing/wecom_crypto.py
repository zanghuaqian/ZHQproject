"""企业微信回调加解密（官方 WXBizMsgCrypt 算法）。"""
from __future__ import annotations

import base64
import hashlib
import socket
import struct
import time
import xml.etree.ElementTree as ET


class WecomCryptoError(Exception):
    pass


def _pkcs7_pad(data: bytes, block=32) -> bytes:
    n = block - (len(data) % block)
    return data + bytes([n] * n)


def _pkcs7_unpad(data: bytes) -> bytes:
    n = data[-1]
    if n < 1 or n > 32:
        raise WecomCryptoError("pkcs7 padding invalid")
    return data[:-n]


class WecomCrypto:
    def __init__(self, token: str, encoding_aes_key: str, corp_id: str):
        self.token = token
        self.corp_id = corp_id
        key = encoding_aes_key + "="
        self.aes_key = base64.b64decode(key)
        if len(self.aes_key) != 32:
            raise WecomCryptoError("EncodingAESKey 解码后必须是 32 字节")
        self.iv = self.aes_key[:16]

    def _aes(self):
        try:
            from Crypto.Cipher import AES
        except ImportError as exc:
            raise WecomCryptoError("请先安装 pycryptodome：pip install pycryptodome") from exc
        return AES.new(self.aes_key, AES.MODE_CBC, self.iv)

    def sign(self, timestamp: str, nonce: str, encrypt: str) -> str:
        items = sorted([self.token, timestamp, nonce, encrypt])
        return hashlib.sha1("".join(items).encode("utf-8")).hexdigest()

    def decrypt(self, encrypt_b64: str) -> str:
        cipher = self._aes()
        raw = cipher.decrypt(base64.b64decode(encrypt_b64))
        raw = _pkcs7_unpad(raw)
        content = raw[16:]
        xml_len = socket.ntohl(struct.unpack("I", content[:4])[0])
        xml_content = content[4 : 4 + xml_len].decode("utf-8")
        from_id = content[4 + xml_len :].decode("utf-8")
        if from_id != self.corp_id:
            raise WecomCryptoError("CorpID 不匹配")
        return xml_content

    def encrypt(self, plaintext: str) -> str:
        import os

        random16 = os.urandom(16)
        xml = plaintext.encode("utf-8")
        packed = struct.pack("I", socket.htonl(len(xml)))
        payload = random16 + packed + xml + self.corp_id.encode("utf-8")
        cipher = self._aes()
        return base64.b64encode(cipher.encrypt(_pkcs7_pad(payload))).decode("utf-8")

    def verify_url(self, msg_signature: str, timestamp: str, nonce: str, echostr: str) -> str:
        if self.sign(timestamp, nonce, echostr) != msg_signature:
            raise WecomCryptoError("URL 验签失败")
        return self.decrypt(echostr)

    def decrypt_message(self, xml_body: str, msg_signature: str, timestamp: str, nonce: str) -> dict:
        encrypt = ET.fromstring(xml_body).findtext("Encrypt") or ""
        if self.sign(timestamp, nonce, encrypt) != msg_signature:
            raise WecomCryptoError("消息验签失败")
        inner = self.decrypt(encrypt)
        root = ET.fromstring(inner)
        return {child.tag: (child.text or "") for child in root}

    def encrypt_reply(self, reply_xml: str, nonce: str, timestamp: str | None = None) -> str:
        ts = timestamp or str(int(time.time()))
        encrypt = self.encrypt(reply_xml)
        signature = self.sign(ts, nonce, encrypt)
        return (
            "<xml>"
            f"<Encrypt><![CDATA[{encrypt}]]></Encrypt>"
            f"<MsgSignature><![CDATA[{signature}]]></MsgSignature>"
            f"<TimeStamp>{ts}</TimeStamp>"
            f"<Nonce><![CDATA[{nonce}]]></Nonce>"
            "</xml>"
        )
