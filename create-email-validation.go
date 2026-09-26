package main

import (
	"errors"
	"log/slog"
)

var NilEmailConfigError = errors.New("email config is nil")
var NoEmailRecipientsError = errors.New("no email recipients")
var SendWithAttachmentsError = errors.New("automatic send not supported with attachments")

func isValidEmailConfig(config *EmailConfig) error {
	if config == nil {
		slog.Error("received nil config")
		return NilEmailConfigError
	}

	if len(config.Recipients) == 0 {
		slog.Error("no recipients specified")
		return NoEmailRecipientsError
	}

	if config.Send && len(config.Attachments) > 0 {
		slog.Error("attachments not allowed when sending")
		return SendWithAttachmentsError
	}

	return nil
}
