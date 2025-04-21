from spack import *
import os

class Mypackage(CMakePackage):
    homepage = "https://example.com"
    url = "https://example.com/mypackage-1.0.tar.gz"

    version('1.0', sha256='...')

    def setup_build_environment(self, env):
        default_mkl_dir = '/opt/intel/mkl'
        mkl_dir = os.environ.get('MKL_DIR', default_mkl_dir)
        env.set('MKL_DIR', mkl_dir)

class Mypackage(CMakePackage):
    variant('mkl_dir', default='/opt/intel/mkl', description='Path to MKL')

    def setup_build_environment(self, env):
        mkl_dir = self.spec.variants['mkl_dir'].value
        env.set('MKL_DIR', mkl_dir)

def cmake_args(self):
    return [
        f'-DCMAKE_INSTALL_PREFIX={self.prefix}',
        ...
    ]

spack install mypackage       # defaults to Release
spack install mypackage build_type=Debug
spack install mypackage build_type=RelWithDebInfo

def install(self, spec, prefix):
    build_type = self.spec.variants['build_type'].value
    build_dir = os.path.join('build', build_type.lower())  # e.g., build/debug

    cmake_args = [
        f'-DCMAKE_BUILD_TYPE={build_type}',
        f'-DCMAKE_INSTALL_PREFIX={prefix}',
        ...
    ]

    mkdirp(build_dir)
    with working_dir(build_dir):
        cmake('../..', *cmake_args)
        make()
        make('install')


spack install --test=root mypackage
def check(self):
    ctest('--preset', 'release')

def check(self):
    if '+tests' in self.spec:
        ctest('--preset', 'release')
before_script:
  - rm -rf package-repo
  - git clone https://gitlab.com/your/package-repo.git
  - spack repo add package-repo
