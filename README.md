# LunchAlertor
Pop a warning when the lunch can be compromised by yet another meeting

We often eat late after most of people around 13h30-14h.
As a result we tend to often forget about meetings starting at 14h.

This script is supposed to be called in the task scheduler around 11h, to remind us about potential early meetings. This way we can plan to eat earlier !

# Usage
To get alert of meetings starting between 12h00 and 15h00:
`python3 LunchAlertor.py 12:00 15:00`
