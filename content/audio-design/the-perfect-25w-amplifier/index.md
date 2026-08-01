+++
title = "The Perfect 25W Amplifier"
weight = 70
+++

How should it be done?

Base it on the [Special Version schematic](/schematics/evolution-of-the-schematics/).
That should be the starting point. Don't attempt to modify an existing amp —
it should be built from scratch!

Then follow the points below:

1. Change the power supply as follows:
2. Increase the pre-stage voltage from 30 to 35V.
3. Place regulators on the pre-stage voltage supply, allowing 5 more volts to
   drop here, taking the unregulated DC voltage up to 40V.
4. Add a cascode stage on the 3rd amplifier stage.
5. Replace the output transistors with modern Japanese types, e.g. the ...
   *(never specified — my notes don't give a part number)*.
6. Refine the other stages, as shown in the suggested schematics.
7. Design a new PCB layout, following the PCB rules detailed below.
8. Use heavy-gauge wire on all supply lines and loudspeaker outlets. If
   possible, use steel/copper bars between the electrolytics.
9. Keep all wires close to the chassis, glue them onto the metal.
10. Place emitter followers before the 3rd stage.
11. Use shielded wires from the input signal jacks to the board. Use two
    signal wires, one for ground and one for active, with the shield
    connected only at one end.

**PCB layout rules:**

1. Keep the input stage separated from the output stage.
2. Place series resistors as close to the base/gate of the receiving
   transistor as possible.
3. Keep EVERYTHING SYMMETRICAL.
4. Short leads everywhere.
5. The higher the network impedance, the shorter the leads.
6. Separate the input and output stage at the connection between the 3rd-stage
   emitter followers and the 3rd-stage common-emitter stage.
7. Use THICK traces on all parts of the output stage. KEEP IT SYMMETRICAL.
   (Any asymmetrical trace routes here will cause an imbalance, with a
   corresponding lack of distortion cancellation — you won't want that, will
   you?)
8. Don't use a ground plane, but guard rings may be useful (never tried them,
   though). A ground plane adds capacitance between all traces and ground,
   reducing high-frequency performance. Remember, this is not a radio — the
   signals do NOT depend on RF reflections and things like that. Capacitance
   is MUCH worse!

Then you'll probably have some questions. Before you get in touch, here are
some anticipated ones, answered up front:

**FAQ**

**Can I sell amplifiers based on this schematic?** No — they are intended for
your personal use. And for your close friends, if that helps you finance the
thing. If you are a company and want to make money on this, contact me
[via the about page](/about/) before my lawyers contact you!

**Can you provide PCBs or components?** I can't provide PCBs or mechanics,
but I may have some suitable electronic components — I'll put up a list of
these later. Check back.

**Can I increase the power output?** Well, you can increase it slightly, up
to perhaps 50W. All output stages including the 3rd stage should have a
corresponding voltage increase, and perhaps you should add more output
transistors in parallel. Note, however, that the more transistors you add in
parallel, the more capacitive load you introduce — 3, or max 4, in parallel.

**What quiescent current should I use?** The quiescent current should be
calculated from each transistor's Hfe-versus-Ic curve. It should be set to no
more than slightly less than half the current at which Hfe has its maximum.
As for the lower limit — don't push it too far down. Note also that more
transistors in parallel means more power loss.
