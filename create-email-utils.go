package main

import (
	"log/slog"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func addRecipient(recipients *ole.IDispatch, recipientInfo EmailRecipient) error {
	recipientMethod, err := oleutil.CallMethod(recipients, "Add", recipientInfo.EmailAddress)
	if err != nil {
		slog.Error("CreateEmail: failed to CallMethod Add", "recipientInfo", recipientInfo, "err", err)
		return err
	}

	slog.Debug("CreateEmail: recipient added", "recipientInfo", recipientInfo)

	recipient := recipientMethod.ToIDispatch()
	defer recipient.Release()

	_, err = oleutil.PutProperty(recipient, "Type", recipientInfo.Type.Value())
	if err != nil {
		slog.Error("CreateEmail: failed to PutProperty Type", "recipientInfo", recipientInfo, "err", err)
		return err
	}
	return nil
}

func addAttachment(attachments *ole.IDispatch, attachmentFilePath string) error {
	attachmentAddMethod, err := oleutil.CallMethod(attachments, "Add", attachmentFilePath)
	if err != nil {
		slog.Error("CreateEmail: failed to CallMethod Add", "attachmentFilePath", attachmentFilePath, "err", err)
		return err
	}

	slog.Debug("CreateEmail: attachment added", "attachmentFilePath", attachmentFilePath)
	attachmentAdd := attachmentAddMethod.ToIDispatch()
	defer attachmentAdd.Release()
	return nil
}
