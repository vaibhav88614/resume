#pragma once
// Mesh data model + polymorphic readers created through a Factory.

#include <algorithm>
#include <array>
#include <cmath>
#include <limits>
#include <memory>
#include <random>
#include <stdexcept>
#include <string>
#include <vector>

struct Point3 {
    double x = 0, y = 0, z = 0;
};

struct BoundingBox {
    Point3 min{ std::numeric_limits<double>::max(),
                std::numeric_limits<double>::max(),
                std::numeric_limits<double>::max() };
    Point3 max{ std::numeric_limits<double>::lowest(),
                std::numeric_limits<double>::lowest(),
                std::numeric_limits<double>::lowest() };

    void expand(const Point3& p) {
        min.x = std::min(min.x, p.x); max.x = std::max(max.x, p.x);
        min.y = std::min(min.y, p.y); max.y = std::max(max.y, p.y);
        min.z = std::min(min.z, p.z); max.z = std::max(max.z, p.z);
    }
    void merge(const BoundingBox& o) {
        expand(o.min);
        expand(o.max);
    }
};

class Mesh {
public:
    std::vector<Point3> points;
};

// Abstract reader interface (polymorphism).
class MeshReader {
public:
    virtual ~MeshReader() = default;
    virtual std::unique_ptr<Mesh> read() = 0;
    virtual std::string name() const = 0;
};

// Generates a random cloud of points; stands in for a real file format so the
// project has no external dependencies. Add e.g. an STL/VTK reader next.
class SyntheticReader final : public MeshReader {
public:
    SyntheticReader(std::size_t count, unsigned seed) : count_(count), seed_(seed) {}

    std::unique_ptr<Mesh> read() override {
        auto mesh = std::make_unique<Mesh>();
        mesh->points.reserve(count_);
        std::mt19937_64 rng(seed_);
        std::uniform_real_distribution<double> dist(-100.0, 100.0);
        for (std::size_t i = 0; i < count_; ++i)
            mesh->points.push_back({ dist(rng), dist(rng), dist(rng) });
        return mesh;
    }
    std::string name() const override { return "synthetic"; }

private:
    std::size_t count_;
    unsigned seed_;
};

// Factory pattern.
inline std::unique_ptr<MeshReader> makeReader(const std::string& kind,
                                              std::size_t count, unsigned seed) {
    if (kind == "synthetic") return std::make_unique<SyntheticReader>(count, seed);
    throw std::invalid_argument("unknown reader kind: " + kind);
}
