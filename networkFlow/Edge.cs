using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TJADSZY.networkFlow
{
    internal class Edge
    {
        public int from, to;
        public Edge residual;
        public double flow;
        public double capacity;

        public Edge(int from, int to, double capacity)
        {
            this.from = from;
            this.to = to;
            this.capacity = capacity;
        }

        public bool isResidual()
        {
            return capacity == 0;
        }

        public double remainingCapacity()
        {
            return capacity - flow;
        }

        public void augment(double bottleNeck)
        {
            flow += bottleNeck;
            residual.flow -= bottleNeck;
        }

        public string edgeInfo()
        {
            return String.Format("(flow:{0}, capacity:{1})", flow, capacity);
        }
    }
}
