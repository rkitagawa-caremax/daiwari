import { useCallback, useState } from 'react';

const CLOSED_CONFIRM_DIALOG = Object.freeze({
  isOpen: false,
  message: '',
  onConfirm: null
});

const CLOSED_ALERT_DIALOG = Object.freeze({
  isOpen: false,
  message: '',
  title: '通知',
  closeOnBackdrop: false
});

export const useAppDialogs = () => {
  const [confirmDialog, setConfirmDialog] = useState(CLOSED_CONFIRM_DIALOG);
  const [alertDialog, setAlertDialog] = useState(CLOSED_ALERT_DIALOG);

  const closeConfirm = useCallback(() => {
    setConfirmDialog(CLOSED_CONFIRM_DIALOG);
  }, []);

  const requestConfirm = useCallback((message, action) => {
    setConfirmDialog({
      isOpen: true,
      message,
      onConfirm: async () => {
        await action();
        closeConfirm();
      }
    });
  }, [closeConfirm]);

  const showAlert = useCallback((message, title = '通知', closeOnBackdrop = false) => {
    setAlertDialog({ isOpen: true, message, title, closeOnBackdrop });
  }, []);

  const closeAlert = useCallback(() => {
    setAlertDialog((current) => ({ ...current, isOpen: false }));
  }, []);

  return {
    confirmDialog,
    alertDialog,
    requestConfirm,
    showAlert,
    closeConfirm,
    closeAlert
  };
};
