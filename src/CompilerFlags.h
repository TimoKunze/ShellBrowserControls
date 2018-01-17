// CompilerFlags.h: Definition of pre-processor constants used for enabling/disabling features.

#pragma once

#define ACTIVATE_SECZONE_FEATURES
#define ACTIVATE_CHUNKED_COLUMNENUMERATIONS
#define ACTIVATE_SUBITEMCONTROL_SUPPORT
//#define SYNCHRONOUS_READDIRECTORYCHANGES					// if defined, we'll use synchronous calls and CancelSychronousIO, otherwise we'll use asynchronous calls and an event

/* If defined, we'll use IEnumShellItems instead of IEnumIDList to enumerate shell items (only on
 * Windows 8+). This simplifies some things, but also has one disadvantage: There seems to be no way to
 * have login dialogs displayed automatically if accessing protected namespaces.
 */
//#define USE_IENUMSHELLITEMS_INSTEAD_OF_IENUMIDLIST
