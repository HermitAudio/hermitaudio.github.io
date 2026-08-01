+++
title = "Troubleshooting the Preamplifier"
weight = 70
+++

I will put up schematics and other material on the preamplifier as well.
However, first a few tips, based on real cases I've had in recent years.

**"Scraping" in the volume control**

This is due to DC leakage through the coupling capacitors, out from the RIAA
stage and into the line stage. Replace the RIAA output electrolytics with
2.2 µF polyester, and the line-stage input capacitors with 0.68 µF polyester.
Remove the 100k loading resistors.
