import base64
from email import policy
from email.parser import BytesParser
from unittest.mock import MagicMock

import pytest
from googleapiclient.errors import HttpError

import secret_santa
from secret_santa import MailManager


def make_http_error():
    response = MagicMock(status=500, reason="Internal Server Error")
    return HttpError(response, b"API request failed")


def test_set_mail_and_subject():
    mailer = MailManager()

    mailer.setMailAndSubject("sender@example.com", "Secret Santa")

    assert mailer.m_from == "sender@example.com"
    assert mailer.m_subject == "Secret Santa"


def test_gmail_send_message_builds_and_sends_email(monkeypatch, capsys):
    mailer = MailManager()
    mailer.setMailAndSubject("sender@example.com", "Secret Santa")
    mailer.m_creds = object()
    set_creds = MagicMock(return_value=True)
    monkeypatch.setattr(mailer, "setCreds", set_creds)

    service = MagicMock()
    response = {"id": "message-123"}
    service.users.return_value.messages.return_value.send.return_value.execute.return_value = response
    build = MagicMock(return_value=service)
    monkeypatch.setattr(secret_santa, "build", build)

    result = mailer.gmailSendMessage(
        "recipient@example.com",
        "You have been chosen to give a gift!",
    )

    set_creds.assert_called_once_with()
    build.assert_called_once_with("gmail", "v1", credentials=mailer.m_creds)
    send = service.users.return_value.messages.return_value.send
    send.assert_called_once()
    assert send.call_args.kwargs["userId"] == "me"

    encoded_message = send.call_args.kwargs["body"]["raw"]
    message_bytes = base64.urlsafe_b64decode(encoded_message.encode())
    message = BytesParser(policy=policy.default).parsebytes(message_bytes)
    assert message["To"] == "recipient@example.com"
    assert message["From"] == "sender@example.com"
    assert message["Subject"] == "Secret Santa"
    assert message.get_content().strip() == "You have been chosen to give a gift!"
    assert result == response
    assert "Message Id: message-123" in capsys.readouterr().out


def test_gmail_send_message_requires_credentials(monkeypatch):
    mailer = MailManager()
    set_creds = MagicMock(return_value=False)
    build = MagicMock()
    monkeypatch.setattr(mailer, "setCreds", set_creds)
    monkeypatch.setattr(secret_santa, "build", build)

    with pytest.raises(RuntimeError, match="Could not set the credentials"):
        mailer.gmailSendMessage("recipient@example.com", "Message body")

    set_creds.assert_called_once_with()
    build.assert_not_called()


def test_gmail_send_message_handles_http_error(monkeypatch, capsys):
    mailer = MailManager()
    set_creds = MagicMock(return_value=True)
    monkeypatch.setattr(mailer, "setCreds", set_creds)
    service = MagicMock()
    service.users.return_value.messages.return_value.send.return_value.execute.side_effect = make_http_error()
    monkeypatch.setattr(secret_santa, "build", MagicMock(return_value=service))

    result = mailer.gmailSendMessage("recipient@example.com", "Message body")

    assert result is None
    assert "An error occurred:" in capsys.readouterr().out
    set_creds.assert_called_once_with()


def test_check_api_prints_labels(monkeypatch, capsys):
    mailer = MailManager()
    mailer.m_creds = object()
    set_creds = MagicMock(return_value=True)
    monkeypatch.setattr(mailer, "setCreds", set_creds)
    service = MagicMock()
    service.users.return_value.labels.return_value.list.return_value.execute.return_value = {
        "labels": [{"name": "INBOX"}, {"name": "SENT"}]
    }
    monkeypatch.setattr(secret_santa, "build", MagicMock(return_value=service))

    mailer.checkAPI()

    set_creds.assert_called_once_with()
    service.users.return_value.labels.return_value.list.assert_called_once_with(userId="me")
    assert capsys.readouterr().out.splitlines() == ["Labels:", "INBOX", "SENT"]


def test_check_api_prints_when_no_labels(monkeypatch, capsys):
    mailer = MailManager()
    monkeypatch.setattr(mailer, "setCreds", MagicMock(return_value=True))
    service = MagicMock()
    service.users.return_value.labels.return_value.list.return_value.execute.return_value = {}
    monkeypatch.setattr(secret_santa, "build", MagicMock(return_value=service))

    mailer.checkAPI()

    assert capsys.readouterr().out.strip() == "No labels found."


def test_check_api_requires_credentials(monkeypatch):
    mailer = MailManager()
    set_creds = MagicMock(return_value=False)
    build = MagicMock()
    monkeypatch.setattr(mailer, "setCreds", set_creds)
    monkeypatch.setattr(secret_santa, "build", build)

    with pytest.raises(RuntimeError, match="Could not set the credentials"):
        mailer.checkAPI()

    set_creds.assert_called_once_with()
    build.assert_not_called()


def test_check_api_handles_http_error(monkeypatch, capsys):
    mailer = MailManager()
    set_creds = MagicMock(return_value=True)
    monkeypatch.setattr(mailer, "setCreds", set_creds)
    service = MagicMock()
    service.users.return_value.labels.return_value.list.return_value.execute.side_effect = make_http_error()
    monkeypatch.setattr(secret_santa, "build", MagicMock(return_value=service))

    mailer.checkAPI()

    assert "An error occurred:" in capsys.readouterr().out
    set_creds.assert_called_once_with()
