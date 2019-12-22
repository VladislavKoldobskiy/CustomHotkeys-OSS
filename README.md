# CustomHotkeys
CustomHotkeys is a (simple) autohotkey script I coded to help me save some mouse clicks.   
Now that's a complete abandonware and a prime example of my terrible coding skills, I'm opensourcing this so it might help other CTS engineers at Microsoft, or anybody outside MS (however, for folks outside Microsoft, I imagine this will not have much value). Feel free to send pull requests if this is something you want to improve.

Stuff this script can do:
* control Dell monitor settings - requires Dell Display Manager software, and monitors USB cables plugged in. I needed to quickly change monitor profiles from "very bright" to "somewhat dim" and change window layouts when I need to demonstrate something. DDM offers a number of command line switches so I just used them
* parse clipboard to find something actionable (case IDs, phone numbers, error codes, etc.). Actions then can be triggered by pressing Right Alt + [J/K/L/;] according to context
* replace some text you want with another one - I used this to keep links I need to send out frequently
* a colleague of mine asked to code in RAlt + Volume Down/Up keys to change music tracks
