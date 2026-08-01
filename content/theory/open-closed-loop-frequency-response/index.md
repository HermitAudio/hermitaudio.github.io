+++
title = "Open and Closed Loop Frequency Response"
weight = 20
math = true
+++

The open-loop frequency response is the frequency response of the amplifier
with no feedback — before feedback, or with the feedback network deliberately
broken.

The closed-loop frequency response is the frequency response of the amplifier
with feedback.

These two are closely related. The theoretical closed-loop frequency response
is equal to the open-loop frequency response times the amount of feedback. If
you have 40dB (100 times) of feedback, and an open-loop response of 1kHz, the
closed-loop frequency response is 100kHz.

The formula relating these two is:

$$f_{cl} = \frac{f_{ol}}{1 + A_{ol} D}$$

where $D$ is the feedback factor and $A_{ol}$ is the open-loop gain. The
total denominator expression is what we call feedback.

The open-loop frequency response is determined by the internal compensation
(intended or not) of the amplifier. Many amplifiers are designed with one
stage having a very high output impedance, so the stray capacitance of that
stage's output determines the open-loop frequency response. For integrated
circuits the open-loop frequency response is either specified, or you can see
it graphically as a function of gain — in the latter case, look at the
maximum gain, which means zero feedback.

Just to remind you: the closed-loop gain $A_{cl}$ is related to the open-loop
gain $A_{ol}$ in exactly the same way as the frequency response, although
inversely.

$$A_{cl} = \frac{A_{ol}}{1 + A_{ol} D}$$

At Electrocompaniet the thinking favored a large open-loop bandwidth. This is
also my opinion, but I feel it shouldn't be larger than necessary. There is
always a tradeoff, and if you go for too high an open-loop bandwidth, you
reduce the possible amount of feedback you can have. My thinking is that as
long as the open-loop bandwidth is high enough, you should use the rest of
your gain for feedback. This will give you a more optimal design, because the
overall distortion will be reduced.

**What determines the open-loop bandwidth**

Mostly it is determined by the last voltage-amplification stage. The
collectors of this stage (assuming transistor amplifiers, which these
articles are all about :-)) are connected to the bases of the drivers of the
output stage. The input impedance of these drivers is normally very
nonlinear, and strongly frequency-dependent. This means you can very well get
a major pole here which varies strongly with signal level and the load
(loudspeaker and cables) of the output stage. The solution to this is to
voltage-drive the output stage, thus loading down the amplification stage.
This will also have the effect of pushing up the cutoff frequency at this
point. The EC amplifiers have this pole around 500kHz. The benefits of a
voltage-driven output stage are described elsewhere on this site, including
in an AES paper (not yet ported to this site).
