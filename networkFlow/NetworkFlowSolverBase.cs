using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TJADSZY.networkFlow
{
    internal class NetworkFlowSolverBase
    {
        public static double INF = double.MaxValue / 2;

        public int n, s, t;
        protected int visitedToken = 1;
        protected int[] visited;

        protected bool solved = false;
        protected double maxFlow;
        protected List<Edge>[] graph;

        public NetworkFlowSolverBase(int n, int s, int t)
        {
            this.n = n;
            this.s = s;
            this.t = t;
            initializeEmptyFlowGraph();
            visited = new int[n];
        }

        private void initializeEmptyFlowGraph()
        {
            graph = new List<Edge>[n];

            for (int i = 0; i < n; i++) graph[i] = new List<Edge>();
        }

        public void addEdge(int from, int to, double capacity)
        {
            Edge e1 = new Edge(from, to, capacity);
            Edge e2 = new Edge(to, from, 0);
            e1.residual = e2;
            e2.residual = e1;
            graph[from].Add(e1);
            graph[to].Add(e2);
        }

        public List<Edge>[] getGraph()
        {
            execute();
            return graph;
        }
        public double getMaxFlow()
        {
            execute();
            return maxFlow;
        }
        private void execute()
        {
            if (solved) return;
            solved = true;
            solve();
        }
        public virtual void solve()
        {
            return;
        }
    }
}
