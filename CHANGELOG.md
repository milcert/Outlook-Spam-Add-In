# Changelog

	+ New feature
	- Bug fix
	? Misc

## 10.05.2019 - v1.2.1.0

	+ Ability to forward internal mails even if restricted
	+ Updated to .NET framework 4.7.2
	- GetCustomUI error message when Office id not found
	- Exception handling for event log write entries
	- Typos and translations

## 01.07.2018 - v1.2.0.0

	+ Changed from RibbonDesigner to XML Ribbon
	+ SpeedUp the starting speed due to the design change
	+ Changed the priority to high for spams containing 1-2 links
	+ Add a contextual menu to report one or more items
	+ High priority when no list unsubscribe with a .ch or .li domain
	- FilterInternalMessages when your Digital ID cannot be found

## 26.09.2017 - v1.1.2.0

	+ Windows 10 x64 compatible
	+ Report email will be sent without read or delivery receipe
	+ Compiled with Option Strict to enforce strict data typing
	+ Better saved Spam path handling with IO.Path.Combine
	+ Mailing list unsubscribe options is now extracted from headers
	+ Better error message when an exception occurs
	+ Avoid spam forward when recipients count >100
	+ Avoid internal messages forwarding per default
	+ To, Cc, and filter internal messages are configurable in registry
	+ A copy of the reported email is not saved upon being sent
	- A few translation corrections in German and Italian
	- Exception can occurs with an empty body, fixed
	? For Office2013 and later, should be added in the HKCU DoNotDisableAddinList

## 03.03.2017 - v1.1.1.0

	+ Forward to Spam Team only mails with an empty spam score, or score < 5
	? Spam score regex catcher modified
	+ End of beta, starting public release (25.04.17-27.04.17)

## 15.02.2017 - v1.1.0.0

	+ Title shortened, priority and spam score are added to it
	+ Better link handling in email (http|https|www)
	+ Extensions added in attachment file blacklist (.pdf, .zip, ...)
	+ Display attachment extension and filename in email report body
	- Fix error message whe opening for the first time (thx to R. G.)
	- Fix report email priority when signed

## 01.11.2016 - v1.0.0.0

	+ Initial beta release