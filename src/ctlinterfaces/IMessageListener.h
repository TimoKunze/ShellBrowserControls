//////////////////////////////////////////////////////////////////////
/// \class IMessageListener
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communication between the controlable controls and \c ShellBrowserControls</em>
///
/// This interface allows \c ShellBrowserControls to hook into the message stream of the controls it can be
/// attached to.
///
/// \sa IInternalMessageListener
//////////////////////////////////////////////////////////////////////


#pragma once


class IMessageListener
{
public:
	/// \brief <em>Allows \c ShellBrowserControls to post-process a message</em>
	///
	/// This method is the very last one that processes a received message. It allows \c ShellBrowserControls
	/// to process the message after the control and after \c DefWindowProc.
	///
	/// \param[in] hWnd The window that received the message.
	/// \param[in] message The received message.
	/// \param[in] wParam The received message's \c wParam parameter.
	/// \param[in] lParam The received message's \c lParam parameter.
	/// \param[out] pResult Receives the result of message processing.
	/// \param[in] cookie A value specified by the \c PreMessageFilter method.
	/// \param[in] eatenMessage If \c TRUE, the control has eaten the message, i. e. it has not forwarded
	///            it to the default window procedure.
	///
	/// \return \c S_OK if the listener processed the message; \c S_FALSE if the listener did not process
	///         the message; \c E_NOTIMPL if the listener does not filter any messages; \c E_POINTER if
	///         \c pResult is an illegal pointer.
	///
	/// \sa PreMessageFilter
	virtual HRESULT PostMessageFilter(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, LRESULT* pResult, DWORD cookie, BOOL eatenMessage) = 0;
	/// \brief <em>Allows \c ShellBrowserControls to pre-process a message</em>
	///
	/// This method is the very first one that processes a received message. It allows
	/// \c ShellBrowserControls to process the message before the control.
	///
	/// \param[in] hWnd The window that received the message.
	/// \param[in] message The received message.
	/// \param[in] wParam The received message's \c wParam parameter.
	/// \param[in] lParam The received message's \c lParam parameter.
	/// \param[out] pResult Receives the result of message processing.
	/// \param[out] pCookie Receives a \c DWORD value that is passed to the \c PostMessageFilter method.
	///
	/// \return \c S_OK if the listener processed the message; \c S_FALSE if the listener did not process
	///         the message; \c E_NOTIMPL if the listener does not filter any messages; \c E_POINTER if
	///         \c pResult or \c pCookie is an illegal pointer.
	///
	/// \sa PostMessageFilter
	virtual HRESULT PreMessageFilter(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, LRESULT* pResult, LPDWORD pCookie) = 0;
};     // IMessageListener