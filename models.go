package main

import "log/slog"

type RecipientType int8

//goland:noinspection GoUnusedConst
const (
	OlTo  RecipientType = 1
	OlCC  RecipientType = 2
	OlBCC RecipientType = 3
)

func (rt RecipientType) Value() int {
	return int(rt)
}

type EmailRecipient struct {
	EmailAddress string
	Type         RecipientType
}

func (e EmailRecipient) LogValue() slog.Value {
	return slog.GroupValue(
		slog.String("EmailAddress", e.EmailAddress),
		slog.Int("RecipientType", e.Type.Value()),
	)
}

type EmailRecipients []EmailRecipient

func (ers EmailRecipients) LogValue() slog.Value {
	values := make([]slog.Value, len(ers))
	for _, u := range ers {
		values = append(values, u.LogValue())
	}
	return slog.AnyValue(values)
}

type EmailAttachments []string

type EmailConfig struct {
	Subject     string
	HTMLBody    string
	Recipients  EmailRecipients
	Attachments EmailAttachments
	SaveDraft   bool
	Send        bool
	Display     bool
}

func (e EmailConfig) LogValue() slog.Value {
	return slog.GroupValue(
		slog.String("Subject", e.Subject),
		slog.String("HTMLBody", e.HTMLBody),
		slog.Any("Recipients", e.Recipients),
		slog.Any("Attachments", e.Attachments),
		slog.Bool("SaveDraft", e.SaveDraft),
		slog.Bool("Display", e.Display),
	)
}
