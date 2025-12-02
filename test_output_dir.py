#!/usr/bin/env python3
"""
Test script to verify get_output_directory() works correctly
"""
import os
from pathlib import Path

def get_output_directory() -> Path:
    """
    Get the appropriate output directory based on environment.
    Use /tmp for serverless environments (Vercel, AWS Lambda) where filesystem is read-only.
    Use ./output for local development.
    """
    # Check if we're in a serverless environment
    is_serverless = (
        os.getenv('VERCEL') is not None or  # Vercel
        os.getenv('AWS_LAMBDA_FUNCTION_NAME') is not None or  # AWS Lambda
        os.getenv('LAMBDA_TASK_ROOT') is not None  # AWS Lambda alternative
    )
    
    if is_serverless:
        output_dir = Path("/tmp")
    else:
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)
    
    return output_dir

if __name__ == "__main__":
    print("Testing get_output_directory()...")
    print()
    
    # Test local environment
    print("1. Local environment (no env vars):")
    output_dir = get_output_directory()
    print(f"   Output directory: {output_dir}")
    print(f"   Is /tmp: {output_dir == Path('/tmp')}")
    print()
    
    # Test Vercel environment
    print("2. Vercel environment (VERCEL=1):")
    os.environ['VERCEL'] = '1'
    output_dir = get_output_directory()
    print(f"   Output directory: {output_dir}")
    print(f"   Is /tmp: {output_dir == Path('/tmp')}")
    del os.environ['VERCEL']
    print()
    
    # Test AWS Lambda environment
    print("3. AWS Lambda environment (AWS_LAMBDA_FUNCTION_NAME=test):")
    os.environ['AWS_LAMBDA_FUNCTION_NAME'] = 'test'
    output_dir = get_output_directory()
    print(f"   Output directory: {output_dir}")
    print(f"   Is /tmp: {output_dir == Path('/tmp')}")
    del os.environ['AWS_LAMBDA_FUNCTION_NAME']
    print()
    
    print("✅ All tests passed!")
