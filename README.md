# SFRT Automated Planning

Automated planning scripts for spatially fractionated radiation therapy (SFRT) / lattice radiation therapy. Two treatment planning systems are supported; each lives in its own place.

## RayStation

The RayStation scripts are in [`RayStation/`](RayStation/) in this repository. They generate the lattice target contours, create the plan, and run the staged dose optimization. See [`RayStation/README.md`](RayStation/README.md) for inputs, outputs, and run order.

## Varian Eclipse (ESAPI)

The Eclipse version is maintained in its own repository:

**https://github.com/vengjean/sfrt-esapi**

It is a C# ESAPI script that creates the lattice peak/valley structures, tuning structures, VMAT beams, and optimization objectives from UI selections and local protocol templates. Build, deployment, and configuration instructions are in that repository's README.

## License

MIT, see [LICENSE](LICENSE).
