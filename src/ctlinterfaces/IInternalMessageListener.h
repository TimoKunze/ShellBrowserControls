//////////////////////////////////////////////////////////////////////
/// \class IInternalMessageListener
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Communication between the controlable controls and \c ShellBrowserControls</em>
///
/// This interface allows the controls, that \c ShellBrowserControls can be attached to, to send
/// \c ShellBrowserControls internal data. The communication is based on messages.
///
/// \sa IMessageListener
//////////////////////////////////////////////////////////////////////


#pragma once


class IInternalMessageListener
{
public:
	/// \brief <em>Processes the messages sent by the attached control</em>
	///
	/// \param[in] message The message to process.
	/// \param[in] wParam Used to transfer data with the message.
	/// \param[in] lParam Used to transfer data with the message.
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT ProcessMessage(UINT message, WPARAM wParam, LPARAM lParam) = 0;
};     // IInternalMessageListener