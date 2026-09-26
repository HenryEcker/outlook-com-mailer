package outlook

import (
	"fmt"
	"log/slog"
	"runtime"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func CreateEmail(config *EmailConfig) error {
	err := isValidEmailConfig(config)
	if err != nil {
		return err
	}
	slog.Debug("CreateEmail: config validated", "config", config)

	/*** REQUIRED FOR THREAD SAFETY ***/
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	slog.Debug("LockOSThread to ensure that all COM calls occur from same thread")

	err = ole.CoInitialize(0)
	if err != nil {
		slog.Error("CreateEmail: failed to initialize COM", "err", err)
		return fmt.Errorf("%v. check if a draft is already open", err)
	}
	defer ole.CoUninitialize()
	slog.Debug("CoInitialize COM")

	/*** CONNECT WITH OUTLOOK COM OBJECT (OUTLOOK DESKTOP MUST BE RUNNING!) ***/
	outlookUnknown, err := oleutil.CreateObject("outlook.application")
	if err != nil {
		slog.Error("CreateEmail: failed to CreateObject outlook.application", "err", err)
		return fmt.Errorf("failed to create COM object for outlook; ensure outlook desktop is installed")
	}
	defer outlookUnknown.Release()
	slog.Debug("CreateEmail: COM outlook.application")

	outlook, err := outlookUnknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		slog.Error("CreateEmail: failed to QueryInterface IID_IDispatch", "err", err)
		return err
	}
	defer outlook.Release()
	slog.Debug("CreateEmail: created outlook via QueryInterface IID_IDispatch")

	/*** CREATE NEW EMAIL ITEM ***/
	messageMethod, err := oleutil.CallMethod(outlook, "CreateItem", 0)
	if err != nil {
		slog.Error("CreateEmail: failed to CallMethod CreateItem", "err", err)
		return err
	}
	message := messageMethod.ToIDispatch()
	defer message.Release()
	slog.Debug("CreateEmail: outlook CreateItem")

	/*** GET INSPECTOR TO BE ABLE TO ACCESS EMAIL PROPERTIES (LIKE EMAIL SIGNATURE) ***/
	_, err = oleutil.CallMethod(message, "GetInspector")
	if err != nil {
		slog.Error("CreateEmail: failed to CallMethod GetInspector", "err", err)
		return err
	}

	/*** ADD RECIPIENTS ***/
	recipientsProperty, err := oleutil.GetProperty(message, "Recipients")
	if err != nil {
		slog.Error("CreateEmail: failed to GetProperty Recipients", "err", err)
		return err
	}
	recipientsDispatch := recipientsProperty.ToIDispatch()
	defer recipientsDispatch.Release()

	for _, recipient := range config.Recipients {
		err = addRecipient(recipientsDispatch, recipient)
		if err != nil {
			return err
		}
	}

	// Resolve All (Makes emails appear with names and links to org cards)
	_, err = oleutil.CallMethod(recipientsDispatch, "ResolveAll")
	if err != nil {
		slog.Error("CreateEmail: failed to CallMethod ResolveAll", "err", err)
		return err
	}

	/*** ADD SUBJECT ***/
	_, err = oleutil.PutProperty(message, "Subject", config.Subject)
	if err != nil {
		slog.Error("CreateEmail: failed to PutProperty Subject", "err", err)
		return err
	}

	/*** ADD BODY ***/
	currentHTMLBody, err := oleutil.GetProperty(message, "HTMLBody")
	if err != nil {
		slog.Error("CreateEmail: failed to GetProperty HTMLBody", "err", err)
		return fmt.Errorf("currentHTMLBody GetProperty %v", err)
	}
	slog.Debug("Got Current Body", "currentHTMLBody", currentHTMLBody)
	// Replace the first blank line in the body with the new body
	newBody := strings.Replace(
		currentHTMLBody.Value().(string),
		"<o:p>&nbsp;</o:p>",
		config.HTMLBody,
		1,
	)
	slog.Debug("Updated Body From Config", "newBody", newBody)
	_, err = oleutil.PutProperty(message, "HTMLBody", newBody)
	if err != nil {
		slog.Error("CreateEmail: failed to PutProperty HTMLBody", "err", err)
		return err
	}

	/*** ADD ATTACHMENTS ***/
	attachmentsProperty, err := oleutil.GetProperty(message, "Attachments")
	if err != nil {
		slog.Error("CreateEmail: failed to GetProperty Attachments", "err", err)
		return err
	}
	attachmentsDispatch := attachmentsProperty.ToIDispatch()
	defer attachmentsDispatch.Release()

	for _, attachment := range config.Attachments {
		err = addAttachment(attachmentsDispatch, attachment)
		if err != nil {
			return err
		}
	}

	/*** SAVE DRAFT ***/
	slog.Debug("CreateEmail: Evaluating Save Draft", "SaveDraft", config.SaveDraft)
	if config.SaveDraft {
		_, err = oleutil.CallMethod(message, "Save")
		if err != nil {
			slog.Error("CreateEmail: failed to CallMethod Save", "err", err)
			return fmt.Errorf("save draft %v", err)
		}
	}

	/*** SEND EMAIL ***/
	slog.Debug("CreateEmail: Evaluating Send MailItem", "Send", config.Send)
	if config.Send {
		_, err = oleutil.CallMethod(message, "Send")
		if err != nil {
			slog.Error("CreateEmail: failed to CallMethod Send", "err", err)
			return fmt.Errorf("send %v", err)
		}
	}

	/*** DISPLAY WINDOW TO USER TO SEND OR EDIT ***/
	slog.Debug("CreateEmail: Evaluating Display", "Display", config.Display)
	if config.Display {
		_, err = oleutil.CallMethod(message, "Display", true)
		if err != nil {
			slog.Error("CreateEmail: failed to CallMethod Display", "err", err)
			return err
		}
	}

	slog.Debug("CreateEmail: successfully created email", "config", config)
	return nil
}
